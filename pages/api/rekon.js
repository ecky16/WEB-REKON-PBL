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
        
        // 1. BACA DATA MITRA (Format Baru 2026)
        const wbMitra = XLSX.readFile(fileMitra.filepath || fileMitra.path);
        const wsMitra = wbMitra.Sheets[wbMitra.SheetNames[0]];
        const dataMitra = XLSX.utils.sheet_to_json(wsMitra, { header: 1 });
        
        const dataVolume = new Map();
        dataMitra.forEach((row, idx) => {
          // Data biasanya mulai stabil di baris ke-9 (index 8)
          if (idx >= 8) {
            const designator = (row[1] || "").toString().trim(); // Kolom B (Designator)
            const vol = parseFloat(row[6]) || 0; // Kolom G (Volume)

            if (vol > 0 && (designator.startsWith('M-') || designator.startsWith('J-'))) {
              dataVolume.set(designator, vol);
            }
          }
        });

        // 2. BACA MASTER TELKOM (File yang Mas simpan di public/data)
        const filePathTelkom = path.join(process.cwd(), 'public', 'data', 'BOQ Telkom.xlsx');
        
        // Cek apakah file Master sudah Mas update ke versi 2026 di GitHub
        if (!fs.existsSync(filePathTelkom)) {
            throw new Error("File Master BOQ Telkom.xlsx tidak ditemukan di public/data");
        }

        const wbTelkom = XLSX.readFile(filePathTelkom);
        const wsTelkom = wbTelkom.Sheets[wbTelkom.SheetNames[0]];

        // 3. UPDATE VOLUME KE MASTER
        const range = XLSX.utils.decode_range(wsTelkom['!ref']);
        for (let R = range.s.r; R <= range.e.r; ++R) {
          const cellAddr = XLSX.utils.encode_cell({ r: R, c: 1 }); // Kolom B di Master
          if (!wsTelkom[cellAddr]) continue;

          const desMaster = wsTelkom[cellAddr].v.toString().trim();
          
          if (dataVolume.has(desMaster)) {
            const cellVolAddr = XLSX.utils.encode_cell({ r: R, c: 6 }); // Kolom G di Master
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
