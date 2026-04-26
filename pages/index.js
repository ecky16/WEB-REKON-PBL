import { useState, useRef } from 'react';

export default function Home() {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [loginError, setLoginError] = useState('');

  const [file, setFile] = useState(null);
  const [mode, setMode] = useState('TA'); // Default: BOQ TA (Dengan Jasa)
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState({ msg: '', type: '' });
  const fileInputRef = useRef(null);

  const VALID_USER = "ecky";
  const VALID_PASS = "anjay";

  const handleLogin = (e) => {
    e.preventDefault();
    if (username === VALID_USER && password === VALID_PASS) {
      setIsLoggedIn(true);
      setLoginError('');
    } else {
      setLoginError('❌ Akses ditolak!');
    }
  };

  const handleUpload = async (e) => {
    e.preventDefault();
    if (!file) return setStatus({ msg: "Pilih file dulu, Mas!", type: 'error' });
    
    setLoading(true);
    setStatus({ msg: `Sedang memproses ${mode}...`, type: 'info' });

    const formData = new FormData();
    formData.append('fileMitra', file);
    formData.append('mode', mode); // Kirim mode ke API

    try {
      const res = await fetch('/api/rekon', { method: 'POST', body: formData });
      if (!res.ok) throw new Error("Gagal memproses file.");

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = mode === 'TA' ? 'HASIL_REKON_TA.xlsx' : 'HASIL_REKON_TIF_NON_JASA.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();

      setStatus({ msg: `✅ Sukses! Format ${mode} terdownload.`, type: 'success' });
      setFile(null); 
      if (fileInputRef.current) fileInputRef.current.value = ''; 
    } catch (err) {
      setStatus({ msg: "Error: " + err.message, type: 'error' });
    } finally {
      setLoading(false);
    }
  };

  if (!isLoggedIn) {
    return (
      <div style={{ minHeight: '100vh', backgroundColor: '#0f172a', display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: 'sans-serif' }}>
        <div style={{ backgroundColor: '#1e293b', padding: '40px', borderRadius: '20px', width: '100%', maxWidth: '350px', textAlign: 'center', border: '1px solid #334155' }}>
          <h2 style={{ color: 'white' }}>Masuk Dashboard</h2>
          <form onSubmit={handleLogin}>
            <input type="text" placeholder="Username" value={username} onChange={(e) => setUsername(e.target.value)} style={{ width: '90%', padding: '12px', margin: '10px 0', borderRadius: '8px', backgroundColor: '#0f172a', color: 'white', border: '1px solid #334155' }} />
            <input type="password" placeholder="Password" value={password} onChange={(e) => setPassword(e.target.value)} style={{ width: '90%', padding: '12px', margin: '10px 0', borderRadius: '8px', backgroundColor: '#0f172a', color: 'white', border: '1px solid #334155' }} />
            <button type="submit" style={{ width: '100%', padding: '12px', borderRadius: '8px', border: 'none', backgroundColor: '#0284c7', color: 'white', fontWeight: 'bold', cursor: 'pointer' }}>Masuk</button>
          </form>
          {loginError && <p style={{ color: '#fca5a5', marginTop: '10px' }}>{loginError}</p>}
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: '100vh', backgroundColor: '#0f172a', color: 'white', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', fontFamily: 'sans-serif', padding: '20px' }}>
      <div style={{ backgroundColor: '#1e293b', padding: '40px', borderRadius: '20px', maxWidth: '500px', width: '100%', border: '1px solid #334155' }}>
        <h1 style={{ textAlign: 'center', color: '#38bdf8' }}>🚀 BOQ Converter 2026</h1>
        
        {/* TAB MENU */}
        <div style={{ display: 'flex', gap: '10px', marginBottom: '25px', marginTop: '20px' }}>
          <button onClick={() => setMode('TA')} style={{ flex: 1, padding: '10px', borderRadius: '8px', border: 'none', backgroundColor: mode === 'TA' ? '#0284c7' : '#334155', color: 'white', cursor: 'pointer', fontWeight: 'bold' }}>
            BOQ MITRA to TA<br/><span style={{fontSize: '10px', fontWeight: 'normal'}}>(Dengan Jasa)</span>
          </button>
          <button onClick={() => setMode('TIF')} style={{ flex: 1, padding: '10px', borderRadius: '8px', border: 'none', backgroundColor: mode === 'TIF' ? '#0284c7' : '#334155', color: 'white', cursor: 'pointer', fontWeight: 'bold' }}>
            BOQ TA to TIF<br/><span style={{fontSize: '10px', fontWeight: 'normal'}}>(Tanpa Jasa)</span>
          </button>
        </div>

        <form onSubmit={handleUpload}>
          <div style={{ border: '2px dashed #334155', padding: '30px', borderRadius: '12px', marginBottom: '20px', textAlign: 'center' }}>
            <input type="file" accept=".xlsx" ref={fileInputRef} onChange={(e) => setFile(e.target.files[0])} style={{ fontSize: '14px' }} />
          </div>
          <button type="submit" disabled={loading || !file} style={{ width: '100%', padding: '14px', borderRadius: '10px', border: 'none', backgroundColor: loading ? '#475569' : '#0284c7', color: 'white', fontWeight: 'bold', cursor: 'pointer' }}>
            {loading ? 'Processing...' : `Proses ${mode}`}
          </button>
        </form>

        {status.msg && <div style={{ marginTop: '20px', padding: '12px', borderRadius: '8px', textAlign: 'center', backgroundColor: status.type === 'error' ? '#7f1d1d' : '#064e3b' }}>{status.msg}</div>}
      </div>
    </div>
  );
}
