# 📊 Personal Financial Dashboard (Excel)

Selamat datang di repositori **Personal Financial Dashboard**! Proyek ini adalah sistem pencatatan dan analisis keuangan pribadi (khususnya disesuaikan untuk kebutuhan mahasiswa) yang dibangun sepenuhnya di atas **Microsoft Excel** tanpa menggunakan macro/VBA, melainkan mengandalkan formula dinamis, *conditional formatting*, *data validation*, dan visualisasi grafik bawaan.

Repositori ini menyimpan seluruh riwayat pengembangan lembar kerja (*spreadsheet*) ini dari versi awal hingga versi stabil terakhir.

---

> [!IMPORTANT]
> **Status Proyek: DIHENTIKAN (Deprecated)**  
> Pengembangan aktif proyek ini telah dihentikan secara resmi karena sistem saat ini dinilai sudah sangat stabil untuk kebutuhan dasar. Selain itu, terdapat kendala teknis berupa kompleksitas formula dan keterbatasan skalabilitas dalam Excel jika dibandingkan dengan membangun aplikasi/program khusus (web/mobile) dengan fungsi serupa.

---

## 🚀 Fitur Utama

Dashboard ini dilengkapi dengan berbagai fitur otomatisasi untuk memudahkan pencatatan keuangan harian:

1. **Dashboard Interaktif**: Halaman ringkasan utama yang menampilkan grafik tren pengeluaran mingguan, persentase alokasi dana, saldo akhir digital/tunai, serta pintasan navigasi cepat.
2. **Pencatatan Terpadu (Pendapatan & Pengeluaran)**: Pencatatan transaksi harian terpisah untuk pemasukan dan pengeluaran dengan klasifikasi metode pembayaran (**Digital** vs **Tunai**).
3. **Pilihan Kategori Dinamis**: Dropdown kategori (*Makanan Pokok*, *Jajan*, *Kebutuhan Lain*) dengan menu detail yang berubah secara dinamis berdasarkan kategori utama yang dipilih.
4. **Indikator Anggaran (Progress Bar)**: Indikator visual berbasis grafik sel (*data bar*) untuk melacak sisa anggaran per kategori secara real-time.
5. **Rekomendasi Batas Harian**: Perhitungan sisa hari dalam sebulan secara otomatis yang dipadukan dengan rekomendasi batas pengeluaran maksimal harian agar anggaran bulanan tidak melampaui batas.
6. **Kalender Peta Panas (Calendar Heatmap)**: Representasi visual tingkat pengeluaran harian pada lembar kerja **Akhir Bulan**, menggunakan *conditional formatting* dengan gradasi warna dari hijau (hemat) hingga merah (intensitas pengeluaran tinggi).
7. **Arsitektur Kalkulasi Terpusat**: Seluruh perhitungan rumit dipisahkan ke dalam lembar kerja khusus (`Kalkulasi`) untuk menjaga tampilan operasional tetap bersih dan meningkatkan performa kalkulasi Excel.

---

## 📂 Struktur Repositori & Versi Template

Repositori ini melacak evolusi template dari versi paling sederhana hingga versi penuh fitur:

| File Template | Versi | Tanggal Rilis | Deskripsi / Fokus Perubahan |
| :--- | :---: | :---: | :--- |
| [Template Keuangan 4.0.1.xlsx](Template%20Keuangan%204.0.1.xlsx) | **v4.0.1** | *Terbaru* | Versi stabil final dengan perbaikan kecil pada formula kalkulasi. |
| [Template Keuangan 4.0.xlsx](Template%20Keuangan%204.0.xlsx) | v4.0.0 | 2025-11-09 | Penambahan Lembar Kerja *Akhir Bulan* (Calendar Heatmap) & Lembar Kalkulasi Terpusat. |
| [Template Keuangan 3.1.xlsx](Template%20Keuangan%203.1.xlsx) | v3.1.x | 2025-10-07 | Perbaikan formula penarikan tunai dan pengenalan kategori "Jajan" dengan sistem budgeting baru. |
| [Template Keuangan 3.0.xlsx](Template%20Keuangan%203.0.xlsx) | v3.0.0 | 2025-08-15 | Perombakan tata letak UI/UX secara masif dan penerapan dropdown berwarna. |
| [Template Keuangan 2.5.xlsx](Template%20Keuangan%202.5.xlsx) | v2.5.0 | 2025-08-05 | Penambahan halaman utama Dashboard, tombol pintasan (hyperlink), dan grafik pai terpisah. |
| [Template Keuangan 2.4.xlsx](Template%20Keuangan%202.4.xlsx) | v2.4.0 | 2025-07-31 | Penggabungan data Digital & Tunai ke dalam satu lembar kerja *Pengeluaran* terpadu. |
| [Template Keuangan 2.3.xlsx](Template%20Keuangan%202.3.xlsx) | v2.3.0 | 2025-06-04 | Pemisahan pencatatan uang digital dan tunai serta pembagian transaksi ke lembar kerja detail. |
| [Template Keuangan 2.2.xlsx](Template%20Keuangan%202.2.xlsx) | v2.2.0 | 2025-05-21 | Implementasi dropdown dinamis (dependen) dan penghitungan otomatis pengeluaran mingguan. |
| [Template Keuangan 2.1.xlsx](Template%20Keuangan%202.1.xlsx) | v2.1.x | 2025-05-19 | Pengelompokan data berdasarkan Makanan Pokok vs Kebutuhan Lain. |
| [Template Keuangan 2.0.xlsx](Template%20Keuangan%202.0.xlsx) | v2.0.0 | 2025-05-14 | Otomatisasi sisa hari hingga akhir bulan dan penambahan diagram tren mingguan. |
| [Template keuangan 1.2.xlsx](Template%20keuangan%201.2.xlsx) | v1.2.0 | 2025-04-07 | Penambahan kalkulasi batas anggaran harian maksimal untuk Uang Digital. |
| [Template keuangan 1.1.xlsx](Template%20keuangan%201.1.xlsx) | v1.1.0 | 2025-02-28 | Penambahan kolom rekap pengeluaran mingguan (Minggu 1 s.d. 4). |
| [Template keuangan 1.0.xlsx](Template%20keuangan%201.0.xlsx) | v1.0.0 | 2025-02-16 | Rilis Awal. Struktur tabel horizontal dengan pencatatan manual sederhana. |

> [!NOTE]  
> Detail lengkap riwayat perubahan dapat dilihat pada berkas [Changelog Dashboard Keuangan Asli.txt](Changelog%20Dashboard%20Keuangan%20Asli.txt).

---

## 📊 Panduan Lembar Kerja (Sheets)

Di dalam versi terbaru (**v4.0.1**), terdapat beberapa lembar kerja dengan fungsinya masing-masing:

*   **`Dashboard`**: Halaman navigasi utama yang menyajikan rangkuman finansial visual (grafik pai persentase uang digital/tunai, grafik garis tren mingguan, serta total saldo saat ini).
*   **`Detail Transaksi`**: Area pemantauan sisa anggaran (*Budgeting*) yang dilengkapi dengan *progress bar* visual untuk memantau pengeluaran per kategori agar tidak *overbudget*.
*   **`Pendapatan`**: Form pencatatan seluruh aliran dana masuk lengkap dengan detail kategori dan metode penerimaan.
*   **`Pengeluaran`**: Form pencatatan seluruh aliran transaksi keluar secara harian dengan fitur dropdown kategori yang interaktif.
*   **`Akhir Bulan`**: Lembar kerja analisis penutupan bulan. Di sini terdapat fitur **Calendar Heatmap** untuk memetakan intensitas pengeluaran harian serta rekapitulasi arus kas bersih bulanan.
*   **`Kalkulasi`**: Lembar kerja backend yang menampung rumus-rumus kompleks untuk mengolah data mentah menjadi informasi siap pakai di lembar Dashboard dan Detail.
*   **`Pendukung`**: Berisi tabel data master (seperti daftar kategori, subkategori, metode pembayaran) yang digunakan sebagai referensi untuk *data validation* (dropdown).

---

## 🛠️ Cara Penggunaan

1. **Unduh Berkas Terbaru**: Silakan unduh berkas versi stabil terakhir [Template Keuangan 4.0.1.xlsx](Template%20Keuangan%204.0.1.xlsx).
2. **Atur Saldo Awal & Anggaran**:
    *   Buka lembar kerja **`Detail Transaksi`**.
    *   Masukkan alokasi anggaran bulanan Anda untuk masing-masing metode pembayaran (Tunai / Digital) pada sel anggaran yang disediakan.
3. **Mulai Mencatat Transaksi**:
    *   Setiap kali menerima uang, catat di lembar kerja **`Pendapatan`**.
    *   Setiap kali membelanjakan uang, catat di lembar kerja **`Pengeluaran`**. Pastikan untuk mengisi kolom **Tanggal**, **Kategori**, **Detail** (menggunakan dropdown), serta **Harga**.
4. **Pantau Finansial Anda**:
    *   Lihat lembar kerja **`Dashboard`** secara berkala untuk memantau grafik tren mingguan Anda.
    *   Perhatikan **`Detail Transaksi`** untuk melihat rekomendasi batas pengeluaran harian dan pastikan indikator *progress bar* tidak berubah menjadi merah yang menandakan anggaran telah terlampaui.

---

## 📝 Catatan Teknis & Batasan

Aplikasi ini menggunakan berbagai formula Excel tingkat lanjut seperti `SUMIFS`, `INDEX`, `MATCH`, dan ketergantungan relasional bertingkat pada *Data Validation*. 
Beberapa keterbatasan yang dijumpai dalam pengembangan versi Excel ini:
*   **Kapasitas Input Terbatas**: Jumlah baris dibatasi (maksimal 80 baris pada lembar kerja Pengeluaran versi v4) untuk menjaga performa dan kestabilan formula Excel. Jika memerlukan baris lebih, formula harus disalin ke baris bawahnya secara manual.
*   **Kompleksitas Rumus**: Penambahan fitur baru membutuhkan pembuatan formula backend yang semakin rumit pada lembar kerja **`Kalkulasi`**, meningkatkan risiko kesalahan referensi sel jika lembar kerja diubah strukturnya.
