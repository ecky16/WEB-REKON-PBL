import formidable from 'formidable';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';

export const config = { api: { bodyParser: false } };

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

  const form = formidable({ keepExtensions: true, uploadDir: '/tmp' });

  return new Promise((resolve) => {
    form.parse(req, async (err, fields, files) => {
      const fileMitra = Array.isArray(files.fileMitra) ? files.fileMitra[0] : files.fileMitra;
      
      try {
        const dataVolume = new Map();
        
        // 1. Baca Mitra pakai STREAMING (Hemat RAM)
        const workbookMitra = new ExcelJS.Workbook();
        const streamMitra = fs.createReadStream(fileMitra.filepath || fileMitra.path);
        const reader = new ExcelJS.xlsx.WorkbookReader(streamMitra, {});
        
        for await (const worksheet of reader) {
          for await (const row of worksheet) {
            if (row.number > 7) {
              const mat = row.getCell(3).text.trim();
              const jas = row.getCell(4).text.trim();
              const vol = parseFloat(row.getCell(9).value) || 0;
              if (vol > 0) {
                if (mat && mat.startsWith('M-')) dataVolume.set(mat, vol);
                if (jas && jas.startsWith('J-')) dataVolume.set(jas, vol);
              }
            }
          }
        }

        // 2. Baca Master Telkom
        const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
        const workbookTelkom = new ExcelJS.Workbook();
        await workbookTelkom.xlsx.readFile(filePathTelkom);
        const outSheet = workbookTelkom.worksheets[0];

        // 3. Update Data (Mapping Tetap)
        outSheet.eachRow((row, rowNumber) => {
          row.eachCell(c => { if (c.alignment) c.alignment.wrapText = false; });
          if (rowNumber >= 9 && rowNumber <= 1082) {
            const des = row.getCell(2).text.trim();
            if ((des.startsWith('M-') || des.startsWith('J-')) && dataVolume.has(des)) {
              row.getCell(7).value = dataVolume.get(des);
            }
          }
        });

        // 4. Kirim Hasil
        const buffer = await workbookTelkom.xlsx.writeBuffer();
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=hasil_rekon.xlsx');
        res.send(buffer);
        
        // Hapus file temp
        if (fs.existsSync(fileMitra.filepath)) fs.unlinkSync(fileMitra.filepath);
        resolve();

      } catch (error) {
        console.error("Error Memory:", error.message);
        res.status(500).json({ error: "Memory Full: " + error.message });
        resolve();
      }
    });
  });
}
