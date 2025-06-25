# GetList

**GetList** adalah aplikasi desktop berbasis **VB WinForms** insert to db and copy files to digdat

---

## 🚀 Fitur Utama

- 🔌 **Koneksi Dinamis ke Oracle**
  - User ID, Password, Host, Port, dan Service Name dapat diatur dan disimpan ke local user settings.  
  
- 📊 **Hasil Query dalam Tabel**
  - Menampilkan hasil eksekusi query dalam `DataGridView`.

- 📤 **Ekspor ke Excel**
  - Dukungan ekspor hasil ke file `.xlsx` menggunakan ClosedXML atau EPPlus.

- 💾 **Penyimpanan Konfigurasi**
  - Detail koneksi disimpan dalam `Properties.Settings.Default` (user scope).

---

## ⚙️ How To RUN?

### 1. **Persiapan**
- Install Visual Studio 2022 atau lebih baru.
- Pastikan sudah menginstall NuGet package berikut:
  - `Oracle.ManagedDataAccess`

### 2. **Konfigurasi Koneksi on setting.setting**
- Jalankan aplikasi.
- Masukkan:
  - Oracle User ID
  - Oracle Password
  - Host (contoh: `10.111.1.111`)
  - Port (default: `1521`)
  - Service Name (contoh: `xx`)
- Klik tombol **"Simpan Koneksi"**
