import formidable from 'formidable';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';

export const config = {
  api: {
    bodyParser: false, // Wajib false untuk upload file
  },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const form = formidable({
    keepExtensions: true,
    allowEmptyFiles: false,
  });

  return new Promise((resolve, reject) => {
    form.parse(req, async (err, fields, files) => {
      if (err) {
        console.error("Error parsing form:", err);
        res.status(500).json({ error: "Gagal membaca upload" });
        return resolve();
      }

      // Pastikan mengambil file yang benar dari object files
      const fileMitra = Array.isArray(files.fileMitra) ? files.fileMitra[0] : files.fileMitra;
      
      if (!fileMitra) {
        res.status(400).json({ error: "File tidak ditemukan" });
        return resolve();
      }

      try {
        // Path file Master Telkom
        const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
        
        if (!fs.existsSync(filePathTelkom)) {
          throw new Error("Master BOQ Telkom tidak ditemukan di server");
        }

        const workbookMitra = new ExcelJS.Workbook();
        await workbookMitra.xlsx.readFile(fileMitra.filepath || fileMitra.path);
        const dataVolume = new Map();

        // Ambil data dari Mitra
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

        // Proses ke Master Telkom
        const workbookTelkom = new ExcelJS.Workbook();
        await workbookTelkom.xlsx.readFile(filePathTelkom);
        const outSheet = workbookTelkom.worksheets[0];

        outSheet.eachRow((row, rowNumber) => {
          // Matikan wrap text
          row.eachCell(c => {
            if (!c.alignment) c.alignment = {};
            c.alignment.wrapText = false;
          });

          if (rowNumber >= 9 && rowNumber <= 1082) {
            const des = row.getCell(2).text.trim();
            // Logika ketat: Hanya M- atau J-
            if ((des.startsWith('M-') || des.startsWith('J-')) && dataVolume.has(des)) {
              row.getCell(7).value = dataVolume.get(des);
            }
          }
        });

        // Kirim hasil sebagai download
        const buffer = await workbookTelkom.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=hasil_rekon_pbl.xlsx');
        res.send(buffer);
        resolve();

      } catch (error) {
        console.error("Proses Error:", error.message);
        res.status(500).json({ error: error.message });
        resolve();
      }
    });
  });
}
