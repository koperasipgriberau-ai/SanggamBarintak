# Aplikasi Koperasi PGRI Sanggam Barintak — Tahap 1

Isi: portal publik (`index.html`), ruang kerja pengurus (`admin.html`), aturan keamanan (`firestore.rules`).

## Pasang (sekitar 30 menit)
1. **Firebase Console** → Add project → tambahkan *Web app* → salin konfigurasinya ke `firebase-config.js`.
2. **Authentication** → Sign-in method → aktifkan *Email/Password*. Atur kebijakan kata sandi (minimal 12 karakter) di pengaturan Authentication. Matikan pendaftaran mandiri bila opsinya tersedia.
3. Authentication → Users → *Add user* untuk sekretaris, bendahara, ketua, dan pengawas. Salin **User UID** masing-masing.
4. **Firestore Database** → Create database → tempel isi `firestore.rules` di tab Rules → Publish.
5. Firestore → Start collection `admins` → ID dokumen = UID pengguna, isi `role` (`bendahara`, `sekretaris`, `ketua`, atau `pengawas`) dan `nama`. Hanya sekretaris & bendahara yang bisa mengubah data.
6. **GitHub**: buat repository, unggah semua file, lalu Settings → Pages → Deploy from branch (`main`, folder root).
7. Firebase → Authentication → Settings → Authorized domains → tambahkan `NAMAANDA.github.io`.
8. Di Google Cloud Console → Credentials, batasi API key hanya untuk domain GitHub Pages Anda.
9. Buka `…/admin.html`, masuk, menu **Impor Excel** → pilih `LAPORAN PEMBUKUAN KOPERASI PGRI.xlsx`.

## Struktur data
- Publik: `publik_anggota`, `publik_ringkasan`, `publik_pengurus`
- Tertutup: `anggota`, `simpanan` (satu dokumen per setoran), `bku`, `surat`, `config/shu`, `log`
- Aturan SHU tersimpan di `config/shu`: cadangan 40%; pendidikan & sosial 5%, pengurus & pengawas 10%, SHU anggota 20%, transaksi anggota 15%, simpanan sukarela 10%.
