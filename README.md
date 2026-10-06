# DocMark — Word Add-in

MVP task-pane add-in untuk Microsoft Word.

## Fitur
- Watermark bebas, default **TERBATAS**.
- Header bebas.
- Apply watermark + header.
- Watermark saja.
- Header saja.
- Remove watermark.
- Menerapkan ke primary / first-page / even-page header pada tiap section.

## PENTING UNTUK iPAD
Word iPad tidak dapat menjalankan add-in dari localhost. File `taskpane.html`, `taskpane.js`, dan `taskpane.css` harus di-host di alamat HTTPS.

### Alur instalasi
1. Upload isi folder ini ke hosting HTTPS statis, misalnya GitHub Pages, Netlify, Cloudflare Pages, atau server kantor.
2. Setelah mendapat URL hosting, buka `manifest.xml`.
3. Untuk `SupportUrl` dan `AppDomain`, isi **origin/domain** hosting, misalnya `https://namamu.github.io`.
4. Untuk `SourceLocation`, isi URL lengkap file `taskpane.html`, misalnya `https://namamu.github.io/upacara-tools/taskpane.html`.
5. Sideload `manifest.xml` ke Word iPad melalui komputer.
6. Tutup Word di iPad sepenuhnya, lalu buka lagi.
7. Buka dokumen Word → **Home → Add-ins → See all / My Add-ins → DocMark**.

## Kenapa bukan tombol custom langsung di Ribbon iPad?
Microsoft belum mendukung Add-in Commands / custom Ribbon commands pada Word iPad. Karena itu versi iPad tampil sebagai task pane.

## Catatan watermark
Watermark utama memakai OOXML/VML di header supaya dapat ditempatkan di belakang isi dokumen. Bila build Word tertentu menolak floating VML, kode punya fallback menjadi teks abu-abu besar di header.
