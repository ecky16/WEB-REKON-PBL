import formidable from 'formidable';
import * as XLSX from 'xlsx';
import path from 'path';
import fs from 'fs';

export const config = {
  api: { bodyParser: false },
};

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

  const form = formidable({});

  return new Promise((resolve) => {
    form.parse(req, async (err, fields, files) => {
      try {
        const fileMitra = Array.isArray(files.fileMitra) ? files.fileMitra[0] : files.fileMitra;
        
        // 1. Baca Mitra
        const wbMitra = XLSX.readFile(fileMitra.filepath || fileMitra.path);
        const wsMitra = wbMitra.Sheets[wbMitra.SheetNames[0]];
        const dataMitra = XLSX.utils.sheet_to_json(wsMitra, { header: 1 });
        
        const dataVolume = new Map();
        dataMitra.forEach((row, idx) => {
          if (idx > 7) {
            const mat = (row[2] || "").toString().trim(); // Kolom C
            const jas = (row[3] || "").toString().trim(); // Kolom D
            const vol = parseFloat(row[8]) || 0; // Kolom I

            if (vol > 0) {
              if (mat.startsWith('M-')) dataVolume.set(mat, vol);
              if (jas.startsWith('J-')) dataVolume.set(jas, vol);
            }
          }
        });

        // 2. Baca Master Telkom
        const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
        const wbTelkom = XLSX.readFile(filePathTelkom);
        const wsTelkom = wbTelkom.Sheets[wbTelkom.SheetNames[0]];

        // 3. Update Volume (Hanya Kolom G)
        const range = XLSX.utils.decode_range(wsTelkom['!ref']);
        for (let R = 8; R <= 1081; ++R) {
          const cellAddr = XLSX.utils.encode_cell({ r: R, c: 1 }); // Kolom B
          if (!wsTelkom[cellAddr]) continue;

          const des = wsTelkom[cellAddr].v.toString().trim();
          if ((des.startsWith('M-') || des.startsWith('J-')) && dataVolume.has(des)) {
            const cellVol = XLSX.utils.encode_cell({ r: R, c: 6 }); // Kolom G
            wsTelkom[cellVol] = { t: 'n', v: dataVolume.get(des) };
          }
        }

        // 4. Kirim Hasil
        const buf = XLSX.write(wbTelkom, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=HASIL_REKON_FINAL.xlsx');
        res.status(200).send(buf);
        resolve();

      } catch (error) {
        res.status(500).json({ error: error.message });
        resolve();
      }
    });
  });
}
