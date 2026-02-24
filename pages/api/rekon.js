import formidable from 'formidable';
import * as XLSX from 'xlsx-js-style';

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
        
        // 1. BACA MITRA
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
              if (mat.startsWith('M-')) dataVolume.set(mat, vol);
              if (jas.startsWith('J-')) dataVolume.set(jas, vol);
            }
          }
        });

        // 2. BACA MASTER TELKOM
        const path = require('path');
        const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
        const wbTelkom = XLSX.readFile(filePathTelkom, { cellStyles: true });
        const wsTelkom = wbTelkom.Sheets[wbTelkom.SheetNames[0]];

        // 3. UPDATE VOLUME (G)
        for (let R = 8; R <= 1081; ++R) {
          const addrDes = XLSX.utils.encode_cell({ r: R, c: 1 }); // Kolom B
          if (!wsTelkom[addrDes]) continue;

          const des = wsTelkom[addrDes].v.toString().trim();
          if ((des.startsWith('M-') || des.startsWith('J-')) && dataVolume.has(des)) {
            const addrVol = XLSX.utils.encode_cell({ r: R, c: 6 }); // Kolom G
            
            // Masukkan Angka
            if (!wsTelkom[addrVol]) wsTelkom[addrVol] = { t: 'n', v: 0 };
            wsTelkom[addrVol].v = dataVolume.get(des);
            
            // Tambahkan Style Manual (Border & Alignment) agar mirip Gambar 1
            wsTelkom[addrVol].s = {
              alignment: { wrapText: false, vertical: 'center', horizontal: 'center' },
              border: {
                top: { style: "thin" }, bottom: { style: "thin" },
                left: { style: "thin" }, right: { style: "thin" }
              }
            };
          }
        }

        // 4. GENERATE HASIL
        const buf = XLSX.write(wbTelkom, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=REKON_TELKOM_STYLISH.xlsx');
        res.status(200).send(buf);
        resolve();

      } catch (error) {
        res.status(500).json({ error: error.message });
        resolve();
      }
    });
  });
}
