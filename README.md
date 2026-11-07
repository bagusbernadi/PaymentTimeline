# PaymentTimeline

**Halaman web:** https://bagusbernadi.github.io/PaymentTimeline/

Halaman timeline pembayaran (pinjaman, pembayaran, dan jadwal ke depan) yang dibuat dengan Python dari data Excel.

## Cara pakai
1. Isi `data_hutang.xlsx` (sheet `Pinjaman` dan `Pembayaran`).
2. Jalankan `python generate_timeline.py` (opsi: `python generate_timeline.py file.xlsx -o index.html`).
3. Commit `index.html` yang dihasilkan; halaman web terbarui otomatis lewat GitHub Pages.

Butuh `openpyxl` (`pip install openpyxl`).
