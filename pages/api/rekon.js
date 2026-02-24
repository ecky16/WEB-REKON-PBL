import formidable from 'formidable';
import ExcelJS from 'exceljs';
import path from 'path';
import fs from 'fs';

export const config = { api: { bodyParser: false } };

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

  const form = formidable({ keepExtensions: true });
  
  try {
    const [fields, files] = await form.parse(req);
    const fileMitra = files.fileMitra[0] || files.fileMitra;

    // 1. CEK FILE MASTER (Penting!)
    const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
    if (!fs.existsSync(filePathTelkom)) {
      throw new Error("File 'BOQ Telkom.xlsx' tidak ditemukan di /public/data/");
    }

    const workbookMitra = new ExcelJS.Workbook();
    await workbookMitra.xlsx.readFile(fileMitra.filepath || fileMitra.path);
    const dataVolume = new Map();

    // 2. BACA DATA MITRA
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

    // 3. PROSES KE MASTER TELKOM
    const workbookTelkom = new ExcelJS.Workbook();
    await workbookTelkom.xlsx.readFile(filePathTelkom);
    const outSheet = workbookTelkom.worksheets[0];

    outSheet.eachRow((row, rowNumber) => {
      // Matikan wrap text biar rapi
      row.eachCell(c => {
        if (!c.alignment) c.alignment = {};
        c.alignment.wrapText = false;
      });

      if (rowNumber >= 9 && rowNumber <= 1082) {
        const des = row.getCell(2).text.trim();
        if ((des.startsWith('M-') || des.startsWith('J-')) && dataVolume.has(des)) {
          row.getCell(7).value = dataVolume.get(des);
        }
      }
    });

    // 4. KIRIM HASIL
    const buffer = await workbookTelkom.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=hasil_rekon.xlsx');
    return res.send(buffer);

  } catch (err) {
    console.error("Detail Error:", err.message);
    return res.status(500).json({ error: err.message });
  }
}
