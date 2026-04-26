import formidable from 'formidable';
import * as XLSX from 'xlsx';
import path from 'path';
import fs from 'fs';

export const config = { api: { bodyParser: false } };

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

  const form = formidable({});

  return new Promise((resolve) => {
    form.parse(req, async (err, fields, files) => {
      try {
        const fileMitra = Array.isArray(files.fileMitra) ? files.fileMitra[0] : files.fileMitra;
        const mode = fields.mode; // Ambil mode dari frontend (TA atau TIF)
        
        const wbMitra = XLSX.readFile(fileMitra.filepath || fileMitra.path);
        const wsMitra = wbMitra.Sheets[wbMitra.SheetNames[0]];
        const dataMitra = XLSX.utils.sheet_to_json(wsMitra, { header: 1 });
        
        const dataVolume = new Map();
        dataMitra.forEach((row, idx) => {
          if (idx > 7) {
            const mat = (row[2] || "").toString().trim(); 
            const jas = (row[3] || "").toString().trim(); 
            const vol = parseFloat(row[8]) || 0; 

            if (vol > 0) {
              // Jika mode TIF, abaikan Jasa
              if (mat.startsWith('M-')) dataVolume.set(mat, vol);
              if (mode === 'TA' && jas.startsWith('J-')) dataVolume.set(jas, vol);
            }
          }
        });

        const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
        const wbTelkom = XLSX.readFile(filePathTelkom);
        const wsTelkom = wbTelkom.Sheets[wbTelkom.SheetNames[0]];

        const range = XLSX.utils.decode_range(wsTelkom['!ref']);
        for (let R = 8; R <= range.e.r; ++R) {
          const cellAddr = XLSX.utils.encode_cell({ r: R, c: 1 });
          if (!wsTelkom[cellAddr]) continue;

          const desMaster = wsTelkom[cellAddr].v.toString().trim();
          const cellVolAddr = XLSX.utils.encode_cell({ r: R, c: 6 });

          // LOGIKA UTAMA:
          if (dataVolume.has(desMaster)) {
            // Jika ada datanya di Mitra, isi volumenya
            wsTelkom[cellVolAddr] = { t: 'n', v: dataVolume.get(desMaster) };
          } else if (mode === 'TIF' && desMaster.startsWith('J-')) {
            // KHUSUS TIF: Jika designator diawali 'J-', pastikan volumenya KOSONG/0
            wsTelkom[cellVolAddr] = { t: 'n', v: 0 };
          }
        }

        const buf = XLSX.write(wbTelkom, { type: 'buffer', bookType: 'xlsx' });
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=REKON_${mode}.xlsx`);
        res.status(200).send(buf);
        resolve();
      } catch (error) {
        res.status(500).json({ error: error.message });
        resolve();
      }
    });
  });
}
