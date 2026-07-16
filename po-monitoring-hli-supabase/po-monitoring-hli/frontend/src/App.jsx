import React from 'react';

const NEW_URL = 'https://serveone.pythonanywhere.com/';

const App = () => {
  return (
    <div
      style={{
        minHeight: '100vh',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        gap: '16px',
        fontFamily: 'sans-serif',
        color: '#1f2937',
        background: '#f9fafb',
        textAlign: 'center',
        padding: '24px',
      }}
    >
      <div
        style={{
          width: 64,
          height: 64,
          borderRadius: '50%',
          background: '#dbeafe',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          fontSize: 28,
        }}
      >
        🚧
      </div>

      <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0 }}>
        Website ini sudah tidak digunakan
      </h1>

      <p style={{ fontSize: 15, color: '#4b5563', maxWidth: 420, margin: 0 }}>
        Serveone Dashboard sudah pindah ke alamat baru. Silakan gunakan link di bawah ini untuk mengakses dashboard.
      </p>

      <a
        href={NEW_URL}
        style={{
          marginTop: 8,
          padding: '12px 24px',
          background: '#2563eb',
          color: '#fff',
          borderRadius: 10,
          fontWeight: 600,
          fontSize: 14,
          textDecoration: 'none',
        }}
      >
        Buka Website Baru
      </a>

      <p style={{ fontSize: 12, color: '#9ca3af', marginTop: 4 }}>
        {NEW_URL}
      </p>
    </div>
  );
};

export default App;
