import { useState, useRef } from 'react';

export default function Home() {
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState({ msg: '', type: '' });
  const fileInputRef = useRef(null);

  const handleUpload = async (e) => {
    e.preventDefault();
    if (!file) return setStatus({ msg: "Pilih file dulu, Mas!", type: 'error' });
    
    setLoading(true);
    setStatus({ msg: 'Sedang memproses data rekon...', type: 'info' });

    const formData = new FormData();
    formData.append('fileMitra', file);

    try {
      const res = await fetch('/api/rekon', { method: 'POST', body: formData });
      if (!res.ok) throw new Error("Gagal memproses file di server.");

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'HASIL_REKON_FINAL.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();

      // --- LOGIKA RESET ---
      setStatus({ msg: '✅ Berhasil! File sudah terdownload.', type: 'success' });
      setFile(null); // Reset state file
      if (fileInputRef.current) fileInputRef.current.value = ''; // Reset tampilan input file
      
    } catch (err) {
      setStatus({ msg: "Waduh error: " + err.message, type: 'error' });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ 
      minHeight: '100vh', 
      backgroundColor: '#0f172a', 
      color: '#f8fafc', 
      display: 'flex', 
      flexDirection: 'column', 
      alignItems: 'center', 
      justifyContent: 'center', 
      fontFamily: 'system-ui, sans-serif',
      padding: '20px' 
    }}>
      <div style={{ 
        backgroundColor: '#1e293b', 
        padding: '40px', 
        borderRadius: '20px', 
        boxShadow: '0 10px 25px rgba(0,0,0,0.3)', 
        maxWidth: '450px', 
        width: '100%',
        textAlign: 'center',
        border: '1px solid #334155'
      }}>
        <h1 style={{ fontSize: '28px', marginBottom: '10px', color: '#38bdf8' }}>🚀 Web Rekon PBL</h1>
        <p style={{ color: '#94a3b8', fontSize: '14px', marginBottom: '30px' }}>
          Konversi otomatis BOQ Mitra ke BOQ Telkom
        </p>

        <form onSubmit={handleUpload}>
          <div style={{
            border: '2px dashed #334155',
            padding: '30px 20px',
            borderRadius: '12px',
            marginBottom: '20px',
            backgroundColor: '#0f172a'
          }}>
            <input 
              type="file" 
              accept=".xlsx" 
              ref={fileInputRef}
              onChange={(e) => {
                setFile(e.target.files[0]);
                setStatus({ msg: '', type: '' });
              }} 
              style={{ fontSize: '14px', width: '100%' }}
            />
          </div>

          <button 
            type="submit" 
            disabled={loading || !file} 
            style={{ 
              width: '100%',
              padding: '14px', 
              borderRadius: '10px',
              border: 'none',
              backgroundColor: loading || !file ? '#475569' : '#0284c7',
              color: 'white',
              fontWeight: 'bold',
              fontSize: '16px',
              cursor: loading || !file ? 'not-allowed' : 'pointer',
              transition: '0.3s'
            }}
          >
            {loading ? 'Processing...' : 'Mulai Rekon!'}
          </button>
        </form>

        {status.msg && (
          <div style={{ 
            marginTop: '25px', 
            padding: '12px', 
            borderRadius: '8px', 
            fontSize: '14px',
            backgroundColor: status.type === 'error' ? '#7f1d1d' : status.type === 'success' ? '#064e3b' : '#0c4a6e',
            color: status.type === 'error' ? '#fca5a5' : status.type === 'success' ? '#6ee7b7' : '#7dd3fc',
            border: `1px solid ${status.type === 'error' ? '#b91c1c' : status.type === 'success' ? '#059669' : '#0369a1'}`
          }}>
            {status.msg}
          </div>
        )}
      </div>

      <div style={{ marginTop: '30px', color: '#475569', fontSize: '12px', textAlign: 'center' }}>
        <p>Dashboard Rekon v2.0 • 2026</p>
        <p>Created with 🔥 for Mas Ecky</p>
      </div>
    </div>
  );
}
