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
        
        // 1. BACA DATA MITRA (Kembali ke Logika Asli yang Sukses)
        const wbMitra = XLSX.readFile(fileMitra.filepath || fileMitra.path);
        const wsMitra = wbMitra.Sheets[wbMitra.SheetNames[0]];
        const dataMitra = XLSX.utils.sheet_to_json(wsMitra, { header: 1 });
        
        const dataVolume = new Map();
        dataMitra.forEach((row, idx) => {
          if (idx > 7) {
            // Kolom C (Material), Kolom D (Jasa), Kolom I (Volume)
            const mat = (row[2] || "").toString().trim(); 
            const jas = (row[3] || "").toString().trim(); 
            const vol = parseFloat(row[8]) || 0; 

            if (vol > 0) {
              if (mat.startsWith('M-')) dataVolume.set(mat, vol);
              if (jas.startsWith('J-')) dataVolume.set(jas, vol);
            }
          }
        });

        // 2. BACA MASTER TELKOM 2026
        const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
        const wbTelkom = XLSX.readFile(filePathTelkom);
        const wsTelkom = wbTelkom.Sheets[wbTelkom.SheetNames[0]];

        // 3. UPDATE VOLUME KE MASTER 2026
        const range = XLSX.utils.decode_range(wsTelkom['!ref']);
        for (let R = 8; R <= range.e.r; ++R) { // Mulai dari baris 9
          const cellAddr = XLSX.utils.encode_cell({ r: R, c: 1 }); // Kolom B di Master
          if (!wsTelkom[cellAddr]) continue;

          const desMaster = wsTelkom[cellAddr].v.toString().trim();
          
          // Hanya isi ke Master jika Designatornya ada di file Mitra
          if ((desMaster.startsWith('M-') || desMaster.startsWith('J-')) && dataVolume.has(desMaster)) {
            const cellVolAddr = XLSX.utils.encode_cell({ r: R, c: 6 }); // Kolom G (Volume) di Master
            wsTelkom[cellVolAddr] = { t: 'n', v: dataVolume.get(desMaster) };
          }
        }

        // 4. KIRIM HASIL
        const buf = XLSX.write(wbTelkom, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=HASIL_REKON_2026.xlsx');
        res.status(200).send(buf);
        resolve();

      } catch (error) {
        console.error("Error Rekon 2026:", error.message);
        res.status(500).json({ error: error.message });
        resolve();
      }
    });
  });
}
