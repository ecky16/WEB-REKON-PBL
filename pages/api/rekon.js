import formidable from 'formidable';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';

export const config = {
  api: { bodyParser: false },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

  const form = formidable({ keepExtensions: true, uploadDir: '/tmp' });

  return new Promise((resolve) => {
    form.parse(req, async (err, fields, files) => {
      const fileMitra = Array.isArray(files.fileMitra) ? files.fileMitra[0] : files.fileMitra;
      
      try {
        const dataVolume = new Map();
        
        // 1. Baca Mitra (Stream agar hemat RAM)
        const workbookMitra = new ExcelJS.Workbook();
        await workbookMitra.xlsx.readFile(fileMitra.filepath || fileMitra.path);
        workbookMitra.worksheets[0].eachRow((row, rowNumber) => {
          if (rowNumber > 7) {
            const mat = row.getCell(3).text.trim();
            const jas = row.getCell(4).text.trim();
            const vol = parseFloat(row.getCell(9).value) || 0;
            if (vol > 0) {
              if (mat.startsWith('M-')) dataVolume.set(mat, vol);
              if (jas.startsWith('J-')) dataVolume.set(jas, vol);
            }
          }
        });

        // 2. Buka Master Telkom
        const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
        const workbookTelkom = new ExcelJS.Workbook();
        await workbookTelkom.xlsx.readFile(filePathTelkom);
        const outSheet = workbookTelkom.worksheets[0];

        // 3. Update Volume & PERTAHANKAN STYLE (Warna/Border)
        outSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          // Matikan Wrap Text tapi pertahankan Fill (Warna) dan Border
          row.eachCell({ includeEmpty: true }, (cell) => {
            if (cell.alignment) {
              cell.alignment = { ...cell.alignment, wrapText: false };
            } else {
              cell.alignment = { wrapText: false };
            }
          });

          // Logika Isi Volume
          if (rowNumber >= 9 && rowNumber <= 1082) {
            const designatorTelkom = row.getCell(2).text.trim();
            if ((designatorTelkom.startsWith('M-') || designatorTelkom.startsWith('J-')) && 
                dataVolume.has(designatorTelkom)) {
              
              const targetCell = row.getCell(7); // Kolom G
              targetCell.value = dataVolume.get(designatorTelkom);
              // Kita tidak menyentuh cell.style agar warna kuning/merah bawaan template tetap ada
            }
          }
        });

        // 4. Kirim Hasil
        const buffer = await workbookTelkom.xlsx.writeBuffer();
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=HASIL_REKON_BERWARNA.xlsx');
        res.status(200).send(buffer);

        // Hapus file temp
        if (fs.existsSync(fileMitra.filepath)) fs.unlinkSync(fileMitra.filepath);
        resolve();

      } catch (error) {
        console.error("Error:", error.message);
        res.status(500).json({ error: error.message });
        resolve();
      }
    });
  });
}
