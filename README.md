# Google Sheets Ticketing API

## Struktur Google Sheets

Nama sheet harus `Orders`. Jalankan `setupSheet()` sekali; script akan membuat header berikut secara otomatis:

| Kolom | Isi |
|---|---|
| Timestamp | Waktu order dibuat |
| Order ID | `GIGS-YYYYMMDD-XXXX` |
| Nama Lengkap | Nama pembeli |
| Email | Email pembeli |
| Nomor WhatsApp | Nomor WhatsApp pembeli |
| Kategori Tiket | Harus sama dengan kategori di `CONFIG.TICKETS` |
| Jumlah Tiket | 1 sampai 10 |
| Total Harga | Harga kategori x jumlah tiket |
| Status Pembayaran | `PENDING` saat order dibuat, dapat diubah ke `SUCCESS` |
| E-Ticket Code / QR Code Unique String | Kode unik untuk e-ticket |
| Metode Pembayaran | `QRIS` atau `TRANSFER_BANK` |
| Bukti Pembayaran | URL file bukti yang disimpan di Google Drive |

Harga dan kuota ada di bagian `CONFIG` pada `Code.gs` dan menjadi sumber kebenaran backend. Saat ini tersedia satu kategori `Presale` seharga Rp60.000. Ubah nilainya sebelum menerima transaksi bila diperlukan. Rekening transfer yang digunakan: BCA `3850962005` a.n. `MUHAMMAD ZAIDAN`.

Gambar QRIS website menggunakan file `qris.jpeg` yang berada satu folder dengan `index.html`.

Bukti pembayaran diunggah otomatis ke folder Google Drive bernama `Payment Proofs`; folder dibuat otomatis jika belum ada. Saat pertama kali menerima order, Apps Script akan meminta izin Google Drive tambahan.

## Setup dan deployment

1. Buat Google Sheet baru.
2. Buka **Extensions > Apps Script**.
3. Tempel isi `Code.gs` dari repository ini.
4. Pastikan `CONFIG.TICKETS`, `PAYMENT_INSTRUCTION`, dan timezone sudah sesuai.
5. Simpan project.
6. Pilih fungsi `setupSheet` pada dropdown fungsi, lalu klik **Run** dan setujui permission.
7. Klik **Deploy > New deployment**.
8. Pilih tipe **Web app**.
9. Atur **Execute as** menjadi **Me**.
10. Atur **Who has access** menjadi **Anyone**.
11. Klik **Deploy**, lalu salin URL yang berakhiran `/exec`.

Jangan gunakan URL `/dev` di website production. Setelah mengubah kode, buat deployment version baru atau gunakan **Deploy > Manage deployments > Edit > New version**.

## Endpoint cek stok

```text
GET https://script.google.com/macros/s/DEPLOYMENT_ID/exec?action=stock
```

Contoh respons:

```json
{
  "status": "success",
  "data": {
    "Presale": { "price": 60000, "quota": 500, "terjual": 2, "remaining": 498 }
  }
}
```

Order berstatus `PENDING` ikut mengurangi stok agar dua pembeli tidak mengambil kuota yang sama. Jika order dibatalkan, ubah statusnya menjadi `CANCELLED` supaya kuota kembali. Jika ingin hanya order `SUCCESS` yang mengurangi stok, ubah `RESERVE_PENDING_STOCK` menjadi `false`.

## Endpoint membuat order

```text
POST https://script.google.com/macros/s/DEPLOYMENT_ID/exec
Content-Type: text/plain;charset=utf-8
```

Payload JSON:

```json
{
  "nama": "Budi Santoso",
  "email": "budi@example.com",
  "wa": "08123456789",
  "kategori": "Presale",
  "jumlah": 2
}
```

Respons sukses berisi `status`, `orderId`, `totalBayar`, `instruksiPembayaran`, dan `eTicketCode`.

## Contoh frontend

Gunakan `text/plain`, bukan `application/json`. `text/plain` adalah simple request sehingga browser tidak mengirim CORS preflight `OPTIONS`, sementara isi request tetap JSON dan dapat dibaca oleh `doPost`.

```html
<script>
  const API_URL = 'https://script.google.com/macros/s/DEPLOYMENT_ID/exec';

  async function buatPesanan(data) {
    const response = await fetch(API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(data)
    });

    const result = await response.json();
    if (!response.ok || result.status !== 'success') {
      throw new Error(result.message || 'Pesanan gagal dibuat.');
    }
    return result;
  }

  async function tampilkanStok() {
    const response = await fetch(API_URL + '?action=stock');
    const result = await response.json();
    if (result.status !== 'success') throw new Error(result.message);
    return result.data;
  }

  buatPesanan({
    nama: 'Budi Santoso',
    email: 'budi@example.com',
    wa: '08123456789',
    kategori: 'Early Bird',
    jumlah: 2
  }).then(console.log).catch(console.error);
</script>
```

## Catatan CORS dan keamanan

- Web App Apps Script tidak menyediakan kontrol header CORS arbitrary seperti server Express biasa. Pola di atas menghindari preflight dengan request sederhana.
- Jangan memakai `mode: 'no-cors'` jika frontend perlu membaca JSON respons; respons opaque tidak dapat diproses.
- Jika deployment tetap ditolak oleh browser atau membutuhkan header CORS khusus, panggil Apps Script melalui backend/proxy pada custom domain Anda. Proxy tersebut menambahkan `Access-Control-Allow-Origin` untuk domain frontend dan meneruskan request ke URL Apps Script.
- Validasi backend tetap wajib. Jangan mempercayai harga, total, kategori, atau kuota dari frontend.
- URL Web App yang dapat diakses `Anyone` bukan tempat menyimpan data rahasia. Untuk produksi, tambahkan token aplikasi atau autentikasi pada proxy dan pertimbangkan proteksi spam/rate limit.
- QR code dapat dibuat di frontend dari `eTicketCode`, atau diproses melalui layanan QR code setelah pembayaran berstatus `SUCCESS`.
