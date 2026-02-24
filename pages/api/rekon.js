import formidable from 'formidable';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';

export const config = {
  api: {
    bodyParser: false,
  },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  // Gunakan folder /tmp karena Vercel hanya izinkan tulis di sana
  const form = formidable({
    keepExtensions: true,
    uploadDir: '/tmp', 
  });

  return new Promise((resolve) => {
    form.parse(req, async (err, fields, files) => {
      if (err) {
        res.status(500).json({ error: "Gagal parsing form" });
        return resolve();
      }

      const fileMitra = Array.isArray(files.fileMitra) ? files.fileMitra[0] : files.fileMitra;
      if (!fileMitra) {
        res.status(400).json({ error: "File mitra tidak terdeteksi" });
        return resolve();
      }

      try {
        // Lokasi Master Telkom - Kita coba dua kemungkinan path
        const rootPath = process.cwd();
        const filePathTelkom = path.join(rootPath, 'public', 'data', 'BOQ Telkom.xlsx');

        if (!fs.existsSync(filePathTelkom)) {
          // Jika gagal di public, coba cari di folder data root (untuk jaga-jaga)
          throw new Error(`Master tidak ketemu di: ${filePathTelkom}`);
        }

        const workbookMitra = new ExcelJS.Workbook();
        await workbookMitra.xlsx.readFile(fileMitra.filepath || fileMitra.path);
        const dataVolume = new Map();

        // 1. Baca Data Mitra (M- dan J- saja)
        const sheetMitra = workbookMitra.worksheets[0];
        sheetMitra.eachRow((row, rowNumber) => {
          if (rowNumber > 7) {
            const mat = row.getCell(3).text.trim();
            const jas = row.getCell(4).text.trim();
            const vol = parseFloat(row.getCell(9).value) || 0;
            if (vol > 0) {
              if (mat && mat.startsWith('M-')) dataVolume.set(mat, vol);
              if (jas && jas.startsWith('J-')) dataVolume.set(jas, vol);
            }
          }
        });

        // 2. Olah Master Telkom
        const workbookTelkom = new ExcelJS.Workbook();
        await workbookTelkom.xlsx.readFile(filePathTelkom);
        const outSheet = workbookTelkom.worksheets[0];

        outSheet.eachRow((row, rowNumber) => {
          // Matikan wrap text biar rapi sesuai request Mas
          row.eachCell(c => {
            if (!c.alignment) c.alignment = {};
            c.alignment.wrapText = false;
          });

          if (rowNumber >= 9 && rowNumber <= 1082) {
            const des = row.getCell(2).text.trim();
            // Hanya isi jika ada M- atau J-
            if ((des.startsWith('M-') || des.startsWith('J-')) && dataVolume.has(des)) {
              row.getCell(7).value = dataVolume.get(des);
            }
          }
        });

        // 3. Kirim Hasil
        const buffer = await workbookTelkom.xlsx.writeBuffer();
        
        // Hapus file temp mitra agar bersih
        if (fs.existsSync(fileMitra.filepath)) fs.unlinkSync(fileMitra.filepath);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=HASIL_REKON_PBL.xlsx');
        res.status(200).send(buffer);
        resolve();

      } catch (error) {
        console.error("LOG ERROR:", error.message);
        res.status(500).json({ error: error.message });
        resolve();
      }
    });
  });
}
