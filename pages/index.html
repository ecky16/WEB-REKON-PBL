import { useState } from 'react';

export default function Home() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleUpload = async (e) => {
    e.preventDefault();
    if (!file) return alert("Pilih file dulu, Mas!");
    
    setLoading(true);
    const formData = new FormData();
    formData.append('fileMitra', file);

    try {
      const res = await fetch('/api/rekon', { method: 'POST', body: formData });
      if (!res.ok) throw new Error("Gagal Rekon");

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'HASIL_REKON_FINAL.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (err) {
      alert("Waduh error: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: '50px', fontFamily: 'sans-serif', textAlign: 'center' }}>
      <h1>🚀 Gebrakan Rekon PBL</h1>
      <p>Upload file Mitra (.xlsx) untuk diconvert ke Template Telkom</p>
      <form onSubmit={handleUpload} style={{ marginTop: '30px' }}>
        <input type="file" accept=".xlsx" onChange={(e) => setFile(e.target.files[0])} />
        <br /><br />
        <button type="submit" disabled={loading} style={{ padding: '10px 20px', cursor: 'pointer' }}>
          {loading ? 'Sedang Memproses...' : 'Mulai Rekon!'}
        </button>
      </form>
    </div>
  );
}
