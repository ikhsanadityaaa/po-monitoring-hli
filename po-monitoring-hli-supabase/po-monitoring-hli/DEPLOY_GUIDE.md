# 🚀 Panduan Deploy — PO HLI Monitoring Dashboard
## Stack: Supabase (Database) + Render.com (Backend) + Vercel (Frontend)

```
[Browser] ──► [Frontend - Vercel] ──► [Backend Flask - Render] ──► [Database - Supabase]
```

---

## LANGKAH 1 — Buat Database di Supabase (GRATIS, data permanen)

1. Buka **https://supabase.com** → Sign Up (gratis)
2. Klik **"New project"**
   - Name: `po-monitoring-hli`
   - Password: buat password yang kuat (simpan!)
   - Region: **Southeast Asia (Singapore)**
3. Tunggu project dibuat (~1 menit)
4. Pergi ke **Settings → Database → Connection string**
5. Pilih tab **"URI"** → copy string seperti ini:
   ```
   postgresql://postgres.[ref]:[password]@aws-0-ap-southeast-1.pooler.supabase.com:6543/postgres
   ```
6. **Simpan string ini** — akan dipakai di Langkah 3

---

## LANGKAH 2 — Upload Kode ke GitHub

1. Buat akun di **https://github.com** → New repository → `po-monitoring-hli` (public)
2. Upload semua file dari folder ini ke repository tersebut
3. Pastikan struktur folder:
   ```
   po-monitoring-hli/
   ├── backend/
   │   ├── app.py
   │   ├── requirements.txt   ← sudah include psycopg2-binary
   │   ├── runtime.txt
   │   ├── Procfile
   │   └── render.yaml
   └── frontend/
       ├── src/App.jsx        ← sudah pakai VITE_API_URL
       ├── package.json
       ├── vite.config.js
       └── vercel.json
   ```

---

## LANGKAH 3 — Deploy Backend ke Render.com

1. Buka **https://render.com** → Sign up pakai GitHub
2. Klik **"New +"** → **"Web Service"**
3. Connect ke repo GitHub kamu
4. **Settings:**
   | Setting | Value |
   |---------|-------|
   | Root Directory | `backend` |
   | Environment | `Python 3` |
   | Build Command | `pip install -r requirements.txt` |
   | Start Command | `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120` |
   | Plan | **Free** |
5. Scroll ke **Environment Variables** → Add:
   - Key: `DATABASE_URL`
   - Value: *(paste connection string Supabase dari Langkah 1)*
6. Klik **"Create Web Service"** → tunggu **"Live"**
7. **Catat URL backend** — contoh: `https://po-monitoring-hli.onrender.com`

> ⚠️ **Free tier Render:** Service tidur setelah 15 menit idle.
> Request pertama setelah tidur ~30 detik. Data aman di Supabase.

---

## LANGKAH 4 — Deploy Frontend ke Vercel

1. Buka **https://vercel.com** → Sign up pakai GitHub
2. Klik **"Add New Project"** → import repo
3. **Settings:**
   | Setting | Value |
   |---------|-------|
   | Root Directory | `frontend` |
   | Framework Preset | Vite |
   | Build Command | `npm run build` |
   | Output Directory | `dist` |
4. **Environment Variables:**
   - Key: `VITE_API_URL`
   - Value: URL Render dari Langkah 3 *(tanpa slash di akhir!)*
     Contoh: `https://po-monitoring-hli.onrender.com`
5. Klik **Deploy** → dapat URL seperti: `https://po-monitoring.vercel.app`

---

## LANGKAH 5 — Test

1. Buka URL Vercel di browser
2. Upload PO List (file Excel) → data tersimpan di Supabase
3. Upload SMRO → dashboard harus menampilkan data
4. Buka **Supabase → Table Editor** untuk lihat/edit data langsung

---

## 🔄 Update Kode

Setiap push ke GitHub → Render dan Vercel otomatis re-deploy.

---

## Kenapa Supabase vs SQLite?

| | SQLite (sebelumnya) | Supabase PostgreSQL |
|---|---|---|
| Data saat Render restart | ❌ Bisa hilang | ✅ Aman permanen |
| Cold start | 30 detik | 30 detik (tapi DB tetap up) |
| Query speed | OK untuk kecil | Lebih baik untuk banyak data |
| Dashboard data | Tidak ada | ✅ Ada di supabase.com |
| Free tier | 1 GB disk | 500 MB database |

---

## 🆓 Ringkasan Biaya

| Platform | Biaya |
|----------|-------|
| Supabase | Gratis (500 MB DB, unlimited rows) |
| Render | Gratis (750 jam/bulan) |
| Vercel | Gratis (unlimited bandwidth) |
| **Total** | **Rp 0** |

