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
  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

  // Pakai folder /tmp karena cuma ini yang diizinkan Vercel untuk nulis file
  const form = formidable({ keepExtensions: true, uploadDir: '/tmp' });

  return new Promise((resolve) => {
    form.parse(req, async (err, fields, files) => {
      if (err) {
        res.status(500).json({ error: "Gagal upload" });
        return resolve();
      }

      const fileMitra = Array.isArray(files.fileMitra) ? files.fileMitra[0] : files.fileMitra;
      
      try {
        // --- LOGIKA SUKSES MAS ECKY DIMULAI DISINI ---
        const workbookMitra = new ExcelJS.Workbook();
        const workbookTelkom = new ExcelJS.Workbook();

        // 1. Ambil Volume dari Mitra (Hanya M- dan J-)
        await workbookMitra.xlsx.readFile(fileMitra.filepath || fileMitra.path);
        const dataVolume = new Map();
        
        workbookMitra.worksheets[0].eachRow((row, rowNumber) => {
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

        // 2. Buka Template Telkom dari folder public/data
        const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
        await workbookTelkom.xlsx.readFile(filePathTelkom);
        const outSheet = workbookTelkom.worksheets[0];

        // 3. Isi Volume & Atur Tampilan (Logika Rapi Mas Ecky)
        outSheet.eachRow((row, rowNumber) => {
          // Atur tampilan: NO WRAP TEXT
          row.eachCell((cell) => {
            if (!cell.alignment) cell.alignment = {};
            cell.alignment.wrapText = false;
          });

          // Logika pengisian Volume
          if (rowNumber >= 9 && rowNumber <= 1082) {
            const designatorTelkom = row.getCell(2).text.trim();
            
            if ((designatorTelkom.startsWith('M-') || designatorTelkom.startsWith('J-')) && 
                dataVolume.has(designatorTelkom)) {
              
              row.getCell(7).value = dataVolume.get(designatorTelkom);
            }
          }
        });

        // Auto-fit kolom sederhana
        outSheet.columns.forEach(column => {
          column.width = column.width < 15 ? 15 : column.width;
        });

        // --- SIMPAN DAN KIRIM KE USER ---
        const buffer = await workbookTelkom.xlsx.writeBuffer();
        
        // Bersihkan RAM & File Sampah
        if (fs.existsSync(fileMitra.filepath)) fs.unlinkSync(fileMitra.filepath);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=HASIL_REKON_FINAL_RAPI.xlsx');
        res.send(buffer);
        resolve();

      } catch (error) {
        console.error("Gagal: ", error.message);
        res.status(500).json({ error: error.message });
        resolve();
      }
    });
  });
}
