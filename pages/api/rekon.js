import formidable from 'formidable';
import * as XLSX from 'xlsx';

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
        
        // Kita gunakan cellStyles: true agar style asli (Kuning/Merah) tetap terbawa
        const wbTelkom = XLSX.readFile(filePathTelkom, { cellStyles: true, cellNF: true, cellDates: true });
        const wsTelkom = wbTelkom.Sheets[wbTelkom.SheetNames[0]];

        // 3. UPDATE VOLUME (G)
        // Mapping Baris 9 (Index 8) sampai Baris 1082 (Index 1081)
        for (let R = 8; R <= 1081; ++R) {
          const cellDesignator = wsTelkom[XLSX.utils.encode_cell({ r: R, c: 1 })]; // Kolom B
          if (!cellDesignator) continue;

          const des = cellDesignator.v.toString().trim();
          
          // Hanya isi jika M- atau J- sesuai request Mas Ecky
          if ((des.startsWith('M-') || des.startsWith('J-')) && dataVolume.has(des)) {
            const addrVol = XLSX.utils.encode_cell({ r: R, c: 6 }); // Kolom G
            
            // Masukkan nilai sambil mempertahankan format sel yang ada
            if(!wsTelkom[addrVol]) wsTelkom[addrVol] = { t: 'n' };
            wsTelkom[addrVol].v = dataVolume.get(des);
            wsTelkom[addrVol].t = 'n';
          }
        }

        // 4. GENERATE DAN KIRIM
        // 'cellStyles: true' sangat penting disini agar warna header tidak hilang
        const buf = XLSX.write(wbTelkom, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
        
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=HASIL_REKON_TELKOM_RAPI.xlsx');
        res.status(200).send(buf);
        resolve();

      } catch (error) {
        console.error("ERROR:", error.message);
        res.status(500).json({ error: error.message });
        resolve();
      }
    });
  });
}
