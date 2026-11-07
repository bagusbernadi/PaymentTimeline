"""Baca data_hutang.xlsx -> hasilkan index.html (timeline hutang).

Pakai:  python generate_timeline.py [file.xlsx] [-o index.html]
Excel:  sheet "Pinjaman"   : Pemberi | Tanggal | Pokok | Catatan
        sheet "Pembayaran" : Tanggal | Pemberi | Jumlah | Status | Keterangan
        Status = "Lunas" (sudah dibayar) atau "Rencana" (jadwal ke depan)
"""
import argparse
import sys
from datetime import date, datetime
from html import escape
from pathlib import Path

from openpyxl import load_workbook

BULAN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
URUT = {"pinjam": 0, "bayar": 1, "lunas": 2, "telat": 3, "rencana": 3}
LABEL = {
    "pinjam": "Pinjaman diterima",
    "bayar": "Pembayaran",
    "lunas": "Lunas",
    "rencana": "Dijadwalkan",
    "telat": "Lewat jatuh tempo",
}


def rp(n):
    return "Rp" + f"{int(round(n)):,}".replace(",", ".")


def tgl(d):
    return f"{d.day} {BULAN[d.month - 1]} {d.year}"


def to_date(v, ctx):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    sys.exit(f"Tanggal tidak valid di {ctx}: {v!r}")


def to_num(v, ctx):
    try:
        return float(v)
    except (TypeError, ValueError):
        sys.exit(f"Angka tidak valid di {ctx}: {v!r}")


def baca_sheet(wb, nama):
    if nama not in wb.sheetnames:
        sys.exit(f'Sheet "{nama}" tidak ditemukan. Sheet ada: {wb.sheetnames}')
    it = wb[nama].iter_rows(values_only=True)
    head = [str(h).strip() for h in next(it)]
    for i, r in enumerate(it, start=2):
        if all(c is None or str(c).strip() == "" for c in r):
            continue
        d = dict(zip(head, r))
        d["_ctx"] = f"{nama} baris {i}"
        yield d


def muat(path, today):
    wb = load_workbook(path, data_only=True)
    pinjaman, events, warn = {}, [], []

    for r in baca_sheet(wb, "Pinjaman"):
        p = str(r["Pemberi"]).strip()
        pokok = to_num(r["Pokok"], r["_ctx"])
        d = to_date(r["Tanggal"], r["_ctx"])
        pinjaman[p] = {"pokok": pokok, "dibayar": 0.0, "tgl": d}
        events.append(dict(tgl=d, jenis="pinjam", pemberi=p, jumlah=pokok,
                           ket=str(r.get("Catatan") or "")))

    for r in baca_sheet(wb, "Pembayaran"):
        p = str(r["Pemberi"]).strip()
        if p not in pinjaman:
            warn.append(f'{r["_ctx"]}: pemberi "{p}" tidak ada di sheet Pinjaman (dilewati)')
            continue
        d = to_date(r["Tanggal"], r["_ctx"])
        jml = to_num(r["Jumlah"], r["_ctx"])
        status = str(r.get("Status") or "").strip().lower()
        if status not in ("lunas", "rencana"):
            warn.append(f'{r["_ctx"]}: Status "{status}" bukan Lunas/Rencana (dianggap Rencana)')
            status = "rencana"
        jenis = "bayar" if status == "lunas" else ("telat" if d < today else "rencana")
        events.append(dict(tgl=d, jenis=jenis, pemberi=p, jumlah=jml,
                           ket=str(r.get("Keterangan") or "")))

    events.sort(key=lambda e: (e["tgl"], URUT[e["jenis"]]))

    sisa_total = 0.0
    hasil = []
    for e in events:
        p = e["pemberi"]
        if e["jenis"] == "pinjam":
            sisa_total += e["jumlah"]
        elif e["jenis"] == "bayar":
            sisa_total -= e["jumlah"]
            pinjaman[p]["dibayar"] += e["jumlah"]
            if pinjaman[p]["dibayar"] > pinjaman[p]["pokok"]:
                warn.append(f"{p}: total pembayaran melebihi pokok")
        if e["jenis"] in ("pinjam", "bayar"):
            e["sisa"] = sisa_total
        hasil.append(e)
        if e["jenis"] == "bayar" and pinjaman[p]["dibayar"] >= pinjaman[p]["pokok"]:
            hasil.append(dict(tgl=e["tgl"], jenis="lunas", pemberi=p,
                              jumlah=pinjaman[p]["pokok"], ket="Hutang ke pemberi ini selesai"))
    return pinjaman, hasil, warn


CSS = """
:root{--bg:#fff;--card:#f6f5f1;--tx:#1f1e1a;--mu:#6b6a63;--line:#d3d1c7;--ac:#0f6e56;
--pinjam:#888780;--bayar:#0f6e56;--lunas:#3b6d11;--rencana:#ba7517;--telat:#a32d2d}
@media(prefers-color-scheme:dark){:root{--bg:#1c1c1a;--card:#2a2a27;--tx:#f1efe8;--mu:#b4b2a9;--line:#444441;--ac:#5dcaa5;
--pinjam:#b4b2a9;--bayar:#5dcaa5;--lunas:#97c459;--rencana:#ef9f27;--telat:#f09595}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--tx);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:720px;margin:0 auto;padding:20px 16px 48px}
h1{font-size:20px;font-weight:500;margin:0 0 2px}
.mu{color:var(--mu);font-size:13px}
.sum{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:16px 0 10px}
.sum div{background:var(--card);border-radius:10px;padding:10px 12px}
.sum b{display:block;font-size:16px;font-weight:500;margin-top:2px}
.bar{height:8px;background:var(--line);border-radius:4px;overflow:hidden}
.bar i{display:block;height:100%;background:var(--ac)}
.pem{display:grid;gap:8px;margin:16px 0}
.pem>div{background:var(--card);border-radius:10px;padding:10px 12px}
.pem .r{display:flex;justify-content:space-between;gap:8px;margin-bottom:6px;font-size:14px}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin:18px 0 14px}
.chips button{font:inherit;font-size:13px;padding:5px 12px;border-radius:99px;border:1px solid var(--line);background:none;color:var(--tx);cursor:pointer}
.chips button.on{background:var(--ac);border-color:var(--ac);color:var(--bg)}
ul.tl{list-style:none;margin:0;padding:0 0 0 24px;border-left:2px solid var(--line)}
.ev{position:relative;margin:0 0 12px}
.ev::before{content:"";position:absolute;left:-32px;top:12px;width:12px;height:12px;border-radius:50%;background:var(--c);border:2px solid var(--bg);box-sizing:content-box}
.ev .card{background:var(--card);border-left:3px solid var(--c);border-radius:8px;padding:9px 12px}
.ev .top{display:flex;justify-content:space-between;gap:8px;font-weight:500}
.ev .tag{font-size:12px;color:var(--c);font-weight:500}
.ev .amt{white-space:nowrap}
.pinjam{--c:var(--pinjam)}.bayar{--c:var(--bayar)}.lunas{--c:var(--lunas)}
.rencana{--c:var(--rencana)}.telat{--c:var(--telat)}
.rencana .card,.telat .card{border-left-style:dashed}
.now{position:relative;margin:4px 0 14px;color:var(--ac);font-size:13px;font-weight:500}
.now::before{content:"";position:absolute;left:-32px;top:5px;width:12px;height:12px;border-radius:50%;background:var(--ac);border:2px solid var(--bg);box-sizing:content-box}
.warn{background:var(--card);border-left:3px solid var(--telat);padding:8px 12px;border-radius:8px;font-size:13px;margin:12px 0}
@media(max-width:480px){.sum{grid-template-columns:1fr}}
"""

JS = """
const chips=document.querySelectorAll('.chips button');
chips.forEach(b=>b.onclick=()=>{chips.forEach(x=>x.classList.remove('on'));b.classList.add('on');
const p=b.dataset.p;document.querySelectorAll('.ev').forEach(e=>e.hidden=p!=='*'&&e.dataset.p!==p)});
"""


def render(pinjaman, events, warn, today):
    total = sum(p["pokok"] for p in pinjaman.values())
    bayar = sum(p["dibayar"] for p in pinjaman.values())
    sisa = total - bayar
    persen = (bayar / total * 100) if total else 0

    h = [f"<h1>Timeline hutang rumah</h1><div class='mu'>Diperbarui {tgl(today)}</div>"]
    h.append("<div class='sum'>"
             f"<div><span class='mu'>Total pinjaman</span><b>{rp(total)}</b></div>"
             f"<div><span class='mu'>Sudah dibayar</span><b>{rp(bayar)}</b></div>"
             f"<div><span class='mu'>Sisa hutang</span><b>{rp(sisa)}</b></div></div>")
    h.append(f"<div class='bar'><i style='width:{persen:.1f}%'></i></div>"
             f"<div class='mu' style='margin-top:4px'>{persen:.1f}% terbayar</div>")

    for w in warn:
        h.append(f"<div class='warn'>{escape(w)}</div>")

    h.append("<div class='pem'>")
    for nama, p in pinjaman.items():
        pr = (p["dibayar"] / p["pokok"] * 100) if p["pokok"] else 0
        h.append(f"<div><div class='r'><span>{escape(nama)}</span>"
                 f"<span class='mu'>sisa {rp(p['pokok'] - p['dibayar'])} dari {rp(p['pokok'])}</span></div>"
                 f"<div class='bar'><i style='width:{pr:.1f}%'></i></div></div>")
    h.append("</div>")

    h.append("<div class='chips'><button class='on' data-p='*'>Semua</button>")
    for nama in pinjaman:
        h.append(f"<button data-p='{escape(nama, True)}'>{escape(nama)}</button>")
    h.append("</div><ul class='tl'>")

    marker_done = False
    for e in events:
        if not marker_done and e["tgl"] > today:
            h.append(f"<li class='now'>Hari ini · {tgl(today)}</li>")
            marker_done = True
        sub = escape(e["pemberi"]) + (f" · {escape(e['ket'])}" if e["ket"] else "")
        extra = f"<div class='mu'>Sisa hutang: {rp(e['sisa'])}</div>" if "sisa" in e else ""
        h.append(
            f"<li class='ev {e['jenis']}' data-p='{escape(e['pemberi'], True)}'><div class='card'>"
            f"<div class='tag'>{LABEL[e['jenis']]} · {tgl(e['tgl'])}</div>"
            f"<div class='top'><span>{sub}</span><span class='amt'>{rp(e['jumlah'])}</span></div>"
            f"{extra}</div></li>")
    if not marker_done:
        h.append(f"<li class='now'>Hari ini · {tgl(today)}</li>")
    h.append("</ul>")

    return (f"<!doctype html><html lang='id'><head><meta charset='utf-8'>"
            f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
            f"<title>Timeline hutang rumah</title><style>{CSS}</style></head>"
            f"<body><main>{''.join(h)}</main><script>{JS}</script></body></html>")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel", nargs="?", default="data_hutang.xlsx")
    ap.add_argument("-o", "--output", default="index.html")
    ap.add_argument("--today", help="YYYY-MM-DD (untuk uji coba)")
    a = ap.parse_args()
    today = datetime.strptime(a.today, "%Y-%m-%d").date() if a.today else date.today()

    pinjaman, events, warn = muat(a.excel, today)
    Path(a.output).write_text(render(pinjaman, events, warn, today), encoding="utf-8")
    for w in warn:
        print("PERINGATAN:", w)
    print(f"OK: {a.output} ({len(events)} kejadian)")


if __name__ == "__main__":
    main()
