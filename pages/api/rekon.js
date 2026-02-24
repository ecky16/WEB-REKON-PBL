import formidable from 'formidable';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

export const config = { api: { bodyParser: false } };

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

  const form = formidable({});
  const [fields, files] = await form.parse(req);
  const fileMitra = files.fileMitra[0];

  try {
    const workbookMitra = new ExcelJS.Workbook();
    await workbookMitra.xlsx.readFile(fileMitra.filepath);
    const dataVolume = new Map();

    // 1. Baca data Mitra
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

    // 2. Buka Master Telkom (Pastikan file ini ada di folder /public/data/)
    const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
    const workbookTelkom = new ExcelJS.Workbook();
    await workbookTelkom.xlsx.readFile(filePathTelkom);
    const outSheet = workbookTelkom.worksheets[0];

    // 3. Mapping Volume & Matikan Wrap Text
    outSheet.eachRow((row, rowNumber) => {
      row.eachCell(c => { if (c.alignment) c.alignment.wrapText = false; });
      if (rowNumber >= 9 && rowNumber <= 1082) {
        const des = row.getCell(2).text.trim();
        if ((des.startsWith('M-') || des.startsWith('J-')) && dataVolume.has(des)) {
          row.getCell(7).value = dataVolume.get(des);
        }
      }
    });

    // 4. Kirim File & Hapus Jejak di RAM
    const buffer = await workbookTelkom.xlsx.writeBuffer();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=hasil.xlsx');
    res.send(buffer);

    // Otomatis terhapus karena fileMitra.filepath di Vercel bersifat temporary
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
