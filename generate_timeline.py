"""Baca data_hutang.xlsx -> hasilkan index.html (timeline hutang).

Pakai:  python generate_timeline.py [file.xlsx] [-o index.html]

Excel:
  sheet "Pinjaman"   : Pemberi | Tanggal | Jumlah Pinjaman | Catatan
  sheet "Pembayaran" : Tanggal Jatuh Tempo | Pemberi | Jumlah | Tanggal Bayar | Status | Keterangan | Sumber Dana (opsional) | Dibayar Oleh (opsional)
    - isi manual: Tanggal Jatuh Tempo, Pemberi, Jumlah, Tanggal Bayar (kosongkan jika belum dibayar), Sumber Dana (opsional)
    - Status dan Keterangan terisi otomatis (rumus di Excel, dihitung juga oleh script ini)
"""
import argparse
import sys
from datetime import date, datetime
from html import escape
from pathlib import Path

from openpyxl import load_workbook

BULAN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
URUT = {"pinjam": 0, "bayar": 1, "talang": 1, "lunas": 2, "telat": 3, "rencana": 3}
LABEL = {
    "pinjam": "Pinjaman diterima",
    "talang": "Talangan",
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


def kosong(v):
    return v is None or str(v).strip() == ""


def baca_sheet(wb, nama, kolom):
    if nama not in wb.sheetnames:
        sys.exit(f'Sheet "{nama}" tidak ditemukan. Sheet ada: {wb.sheetnames}')
    it = wb[nama].iter_rows(values_only=True)
    head = [str(h).strip() for h in next(it)]
    for i, r in enumerate(it, start=2):
        d = dict(zip(head, r))
        d["_ctx"] = f"{nama} baris {i}"
        # abaikan baris yang kolom isian manualnya kosong semua
        if all(kosong(d.get(k)) for k in kolom):
            continue
        yield d


def muat(path, today):
    wb = load_workbook(path, data_only=True)
    pinjaman, events, warn, tot_key = {}, [], [], {}

    for r in baca_sheet(wb, "Pinjaman", ("Pemberi", "Tanggal", "Jumlah Pinjaman", "Pokok")):
        p = str(r["Pemberi"]).strip()
        nilai = r.get("Jumlah Pinjaman", r.get("Pokok"))
        pokok = to_num(nilai, r["_ctx"])
        d = to_date(r["Tanggal"], r["_ctx"])
        asal = str(r.get("Asal Dana") or "").strip()
        pinjaman.setdefault(p, {"pokok": 0.0, "dibayar": 0.0})
        pinjaman[p]["pokok"] += pokok
        tot_key[(p, asal)] = tot_key.get((p, asal), 0.0) + pokok
        events.append(dict(tgl=d, jenis="pinjam", pemberi=p, jumlah=pokok, asal=asal,
                           ket=str(r.get("Catatan") or "")))

    baris = []
    for r in baca_sheet(wb, "Pembayaran", ("Tanggal Jatuh Tempo", "Pemberi", "Jumlah", "Tanggal Bayar")):
        p = str(r.get("Pemberi") or "").strip()
        if p not in pinjaman:
            warn.append(f'{r["_ctx"]}: pemberi "{p}" tidak ada di sheet Pinjaman (dilewati)')
            continue
        due = to_date(r["Tanggal Jatuh Tempo"], r["_ctx"])
        jml = to_num(r["Jumlah"], r["_ctx"])
        tb = None if kosong(r.get("Tanggal Bayar")) else to_date(r["Tanggal Bayar"], r["_ctx"])
        if tb and tb > today:
            warn.append(f'{r["_ctx"]}: Tanggal Bayar {tgl(tb)} ada di masa depan')
        oleh = str(r.get("Dibayar Oleh") or "").strip()
        if oleh and oleh not in pinjaman:
            warn.append(f'{r["_ctx"]}: "Dibayar Oleh" {oleh} tidak ada di sheet Pinjaman (dianggap dibayar sendiri)')
            oleh = ""
        baris.append(dict(pemberi=p, due=due, jumlah=jml, tb=tb, oleh=oleh,
                          asal=str(r.get("Asal Dana") or "").strip(),
                          src=str(r.get("Sumber Dana") or "").strip()))

    # Total per kunci (pemberi, asal dana) = pinjaman + talangan yang dibayarkan pemberi itu untuk pihak lain
    for x in baris:
        if x["oleh"]:
            tot_key[(x["oleh"], "")] = tot_key.get((x["oleh"], ""), 0.0) + x["jumlah"]

    # Keterangan otomatis: "Cicilan n" atau "Pelunasan" (sama dengan rumus di Excel)
    for b in baris:
        key = (b["pemberi"], b["asal"])
        sama = [x for x in baris if (x["pemberi"], x["asal"]) == key and x["due"] <= b["due"]]
        kum = sum(x["jumlah"] for x in sama)
        tot = tot_key.get(key, 0.0) or 1e99
        b["ket"] = "Pelunasan" if kum >= tot - 0.5 else f"Cicilan {len(sama)}"
        if b["tb"]:
            b["status"] = "Lunas"
        else:
            b["status"] = "Terlambat" if b["due"] < today else "Rencana"
        jenis = {"Lunas": "bayar", "Terlambat": "telat", "Rencana": "rencana"}[b["status"]]
        src = b["src"]
        oleh = b["oleh"] if jenis == "bayar" else ""
        if oleh:
            src = f"Ditalangi {oleh}" + (f" · {src}" if src else "")
        events.append(dict(tgl=b["tb"] or b["due"], jenis=jenis, pemberi=b["pemberi"],
                           jumlah=b["jumlah"], ket=b["ket"], src=src, src0=b["src"], oleh=oleh,
                           asal=b["asal"]))

    events.sort(key=lambda e: (e["tgl"], URUT[e["jenis"]]))

    run = {p: {"pokok": 0.0, "dibayar": 0.0} for p in pinjaman}
    run_a = {}
    sisa = {p: 0.0 for p in pinjaman}
    total_pinjam = bayar_sendiri = 0.0
    hasil = []
    for e in events:
        p, a = e["pemberi"], e.get("asal", "")
        extra = []
        if e["jenis"] == "pinjam":
            run[p]["pokok"] += e["jumlah"]
            sisa[p] += e["jumlah"]
            total_pinjam += e["jumlah"]
            if a:
                run_a.setdefault((p, a), {"pokok": 0.0, "dibayar": 0.0})["pokok"] += e["jumlah"]
        elif e["jenis"] == "bayar":
            run[p]["dibayar"] += e["jumlah"]
            sisa[p] -= e["jumlah"]
            if a and (p, a) in run_a:
                run_a[(p, a)]["dibayar"] += e["jumlah"]
            if e["oleh"]:  # hutang pindah ke penalang, total tidak berubah
                o = e["oleh"]
                run[o]["pokok"] += e["jumlah"]
                sisa[o] += e["jumlah"]
                extra.append(dict(tgl=e["tgl"], jenis="talang", pemberi=o, jumlah=e["jumlah"],
                                  ket=f"Menalangi pembayaran ke {p}", src=e["src0"], sisa=sisa[o]))
            else:
                bayar_sendiri += e["jumlah"]
        if e["jenis"] in ("pinjam", "bayar"):
            e["sisa_main"] = sisa[p]
            e["sisa"] = (run_a[(p, a)]["pokok"] - run_a[(p, a)]["dibayar"]) if a and (p, a) in run_a else sisa[p]
        hasil.append(e)
        hasil.extend(extra)
        if e["jenis"] == "bayar":
            if a and (p, a) in run_a:
                if run_a[(p, a)]["dibayar"] >= run_a[(p, a)]["pokok"] - 0.5:
                    hasil.append(dict(tgl=e["tgl"], jenis="lunas", pemberi=p, asal=a,
                                      jumlah=run_a[(p, a)]["pokok"], ket="Hutang dari pihak ini selesai"))
            elif run[p]["dibayar"] >= run[p]["pokok"] - 0.5:
                hasil.append(dict(tgl=e["tgl"], jenis="lunas", pemberi=p,
                                  jumlah=run[p]["pokok"], ket="Hutang ke pemberi ini selesai"))
    for p in pinjaman:
        pinjaman[p].update(run[p])
        if run[p]["dibayar"] > run[p]["pokok"] + 0.5:
            warn.append(f"{p}: total pembayaran melebihi jumlah pinjaman")
    for k, v in run_a.items():
        pass
    ring = dict(total=total_pinjam, bayar=bayar_sendiri, sisa=total_pinjam - bayar_sendiri)
    return pinjaman, hasil, warn, ring


CSS = """
:root{--bg:#fff;--card:#f4f1e8;--tx:#1b2233;--mu:#5b5f6b;--line:#e1dcca;
--blue:#1e459f;--red:#cf2a2a;--yellow:#fabd32;--beige:#e1dcca}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--tx);font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif}
main{max-width:1400px;margin:0 auto;padding:20px 16px 48px}
h1{font-size:22px;font-weight:600;margin:0 0 2px;letter-spacing:-.01em}
.mu{color:var(--mu);font-size:13px}
.beli{margin:22px 0 8px;border:1px solid var(--beige);border-radius:14px;padding:14px 16px;background:#fff}
.bh2{display:flex;justify-content:space-between;align-items:baseline;font-size:14px;color:var(--tx)}.bh2 b{font-size:18px}
.sbar{display:flex;height:10px;border-radius:5px;overflow:hidden;margin:10px 0 6px;background:var(--beige)}
.sbar i,.leg i{display:block;height:100%}.b1{background:var(--blue)}.b2{background:var(--yellow)}
.leg{display:flex;gap:16px;flex-wrap:wrap;font-size:12px;color:#6b7280}.leg span{display:flex;align-items:center;gap:6px}.leg i{width:10px;height:10px;border-radius:3px}
.bg{display:grid;gap:14px;margin-top:12px}@media(min-width:700px){.bg{grid-template-columns:1fr 1fr}}
.bcol{background:#faf8f2;border-radius:10px;padding:10px 12px}
.bt{display:flex;justify-content:space-between;font-weight:600;font-size:13px;padding-bottom:6px;border-bottom:1px solid var(--beige);color:var(--tx)}
.bi2{display:flex;justify-content:space-between;gap:10px;padding:7px 0 7px 9px;margin-top:6px;border-left:4px solid var(--yellow);font-size:13px;color:var(--tx)}
.bi2.p{border-left-color:var(--blue)}.bi2 b{white-space:nowrap}.ri{font-size:12px;color:#6b7280}
.sum{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:16px 0 12px}
.sum div{background:var(--f);border-top:4px solid var(--c);border-radius:12px;padding:10px 12px}
.sum b{display:block;font-size:16px;font-weight:600;margin-top:2px;color:var(--tx)}
.s-pinjam{--c:var(--blue);--f:#dde5f6;--t:#1e459f}
.s-bayar{--c:var(--yellow);--f:#fdefc8;--t:#7a5600}
.s-sisa{--c:var(--red);--f:#f9dede;--t:#b31f1f}
.bar{height:8px;background:#efebdd;border-radius:4px;overflow:hidden}
.bar i{display:block;height:100%;background:var(--yellow)}
.pem{display:grid;gap:10px;margin:16px 0}
.pem>div{background:var(--card);border-radius:12px;padding:10px 12px}
.pem .r{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:6px;font-size:14px}
.pem .done{background:var(--blue);color:#fff}
.pem .done .mu{color:#dde5f6}
.pem .done .bar{background:rgba(255,255,255,.25)}
.pem .done .bar i{background:var(--yellow)}
.badge{background:var(--yellow);color:var(--tx);border-radius:99px;padding:1px 10px;font-size:12px;font-weight:600;white-space:nowrap}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin:18px 0 14px}
.chips button{font:inherit;font-size:13px;padding:5px 12px;border-radius:99px;border:1px solid var(--beige);background:none;color:var(--tx);cursor:pointer}
.chips button.on{background:var(--blue);border-color:var(--blue);color:#fff}
ul.tl{list-style:none;margin:0;padding:0 0 0 24px;border-left:2px solid var(--beige)}
.ev{position:relative;margin:0 0 12px}
.ev::before{content:"";position:absolute;left:-32px;top:12px;width:12px;height:12px;border-radius:50%;background:var(--c);border:2px solid #fff;box-sizing:content-box}
.ev .card{background:var(--f);border-left:4px solid var(--c);border-radius:12px;padding:9px 12px}
.ev .top{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:flex-start;gap:2px 10px;font-weight:600}
.ev .lf{min-width:0;flex:1 1 150px}.ev .rt{text-align:right;margin-left:auto}
.ev .src{font-weight:400;color:var(--mu);font-size:13px}
.ev .tag{font-size:12px;color:var(--mu);font-weight:600}
.lunas .tag,.telat .tag{color:rgba(255,255,255,.85)}
.ev .amt{white-space:nowrap}
.lns{background:var(--blue);color:#fff;border-radius:99px;padding:1px 9px;font-size:12px;font-weight:600;margin-right:8px}
.ev .sisa{color:#b31f1f;font-size:14px;font-weight:600}
.pinjam{--c:var(--blue);--f:#dde5f6;--t:#1e459f}
.bayar{--c:var(--yellow);--f:#fdefc8;--t:#7a5600}
.rencana{--c:#b8af92;--f:#efebdd;--t:#5c5640}
.telat{--c:var(--red);--f:var(--red);--t:#fff}
.lunas{--c:var(--yellow);--f:var(--blue);--t:#fabd32}
.lunas .card,.telat .card{color:#fff}
.telat .card{border-left-color:#fff}
.rencana .card{border-left-style:dashed}
.now{position:relative;margin:4px 0 14px;color:var(--tx);font-size:13px;font-weight:600}
.now::before{content:"";position:absolute;left:-32px;top:5px;width:12px;height:12px;border-radius:50%;background:var(--tx);border:2px solid #fff;box-sizing:content-box}
.talang{--c:var(--blue);--f:#dde5f6}
.trunk{margin-top:22px}
.tr{width:330px;max-width:100%;background:var(--tx);color:#fff;border-radius:12px;padding:8px 14px;text-align:center;margin:0 auto}
.tr span{display:block;font-size:12px;opacity:.85}.tr b{font-size:17px;font-weight:600}
.stem{width:2px;height:18px;background:var(--beige);margin:0 auto}
.tgrid{display:flex;flex-direction:column;gap:0}
@media(min-width:700px){.tgrid{display:grid;grid-template-columns:84px var(--cols);column-gap:28px}
.tr{margin:0 0 0 calc(var(--m) - 165px)}.stem{margin:0 0 0 calc(var(--m) - 1px)}
.cell>.ev .dt{display:none}}
.norail{border-left-color:transparent!important}
.fseg{list-style:none;position:absolute;top:0;left:-27px;width:25px;border-top:2px solid var(--beige);height:0}
.tgrid ul.tl.cell{position:relative}
.sfc{position:relative}
.fseg{display:none}@media(min-width:700px){.fseg{display:block}}
.fl{display:none;position:absolute;top:0;border-top:2px solid var(--beige)}
@media(min-width:700px){.fl{display:block}}
.sfc .sh{display:flex;flex-direction:column;gap:2px;background:var(--sc);color:#fff;border-radius:10px;padding:7px 12px;margin:0 0 14px;font-size:13px}
.sfc .badge{align-self:flex-start}
.bs2{width:2px;height:14px;background:var(--sc);margin:0 auto}
@media(max-width:699px){.emp,.dg{display:none}}
.dg{font-size:12px;font-weight:600;color:var(--mu);text-align:right;padding-top:12px;white-space:nowrap}
.dg.dnow{color:var(--tx)}
.bl,.bl2{display:none}
@media(min-width:700px){.bl{display:block;position:absolute;top:0;border-top:2px solid var(--beige)}}
.branch{position:relative}

.bs{width:2px;height:18px;background:var(--bc);margin:0 auto}
.bh{background:var(--bc);color:#fff;border-radius:12px;padding:10px 14px;margin-bottom:16px}
.bh .r{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:8px}
.bh .nm{font-weight:600;font-size:16px}.bh .bi{font-size:13px;opacity:.92}
.bbar{height:8px;background:rgba(255,255,255,.28);border-radius:4px;overflow:hidden}
.bbar i{display:block;height:100%;background:#fff}
.tgrid ul.tl.cell{border-left:3px solid var(--bc);margin-left:10px}
.cell .ev::before{left:-33px}.cell .now::before{left:-33px}
.sf{list-style:none;margin:0 0 12px -2px}
.sstem{width:2px;height:16px;background:var(--beige);margin:0 auto}
.sgrid{display:grid;grid-template-columns:1fr;gap:4px 16px}
.sb2{position:relative;min-width:0}
.bs2{width:2px;height:14px;background:var(--sc);margin:0 auto}
.sb2 .sh{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:6px;background:var(--sc);color:#fff;border-radius:10px;padding:6px 12px;margin:0 0 14px;font-size:13px}
.sh .nm{font-weight:600;white-space:nowrap}.sh .bi{opacity:.92;font-size:12px}
.cell .sb2 ul.tl{border-left:3px solid var(--sc);margin-left:10px}
@media(min-width:1100px){.sgrid{grid-template-columns:repeat(var(--n),1fr)}
.bl2{display:block;position:absolute;top:0;border-top:2px solid var(--beige)}}
.warn{background:var(--card);border-left:3px solid var(--red);padding:8px 12px;border-radius:8px;font-size:13px;margin:12px 0}
@media(max-width:480px){.sum{grid-template-columns:1fr}}
"""

JS = """
const chips=document.querySelectorAll('.chips button');
chips.forEach(b=>b.onclick=()=>{chips.forEach(x=>x.classList.remove('on'));b.classList.add('on');
const p=b.dataset.p;document.querySelectorAll('.ev').forEach(e=>e.hidden=p!=='*'&&e.dataset.p!==p)});
"""


WARNA = ["#7b3fa0", "#0f7b6c", "#be185d", "#475569"]
SUBWARNA = ["#0369a1", "#c2410c", "#4d7c0f", "#a21caf"]
G = 28  # px, jarak antar kolom
GUT = 84  # px, lebar kolom tanggal di layar lebar
W_MAIN, W_SUB = 1.0, 0.75


def centro(w, a, b, s):
    """Pusat span kolom a..b (grid fr berbobot w, jarak G, area mulai s px) -> (persen lebar, offset px)."""
    t, n = sum(w), len(w)
    k = sum(w[:a]) + sum(w[a:b + 1]) / 2
    return k / t * 100, s - k * s / t - k * (n - 1) * G / t + (a + (b - a) / 2) * G


def kartu(e):
    judul = escape(e["ket"]) if e["ket"] else LABEL[e["jenis"]]
    src = f"<div class='src'>{escape(e['src'])}</div>" if e.get("src") else ""
    pill = "<span class='lns'>Lunas</span>" if "sisa" in e and e["sisa"] <= 0.5 and e["jenis"] == "bayar" else ""
    sisa_h = f"<div class='sisa'>{pill}{rp(e['sisa'])}</div>" if "sisa" in e else ""
    return (f"<li class='ev {e['jenis']}'><div class='card'>"
            f"<div class='tag'>{LABEL[e['jenis']]}<span class='dt'> · {tgl(e['tgl'])}</span></div>"
            f"<div class='top'><div class='lf'><span>{judul}</span>{src}</div>"
            f"<div class='rt'><div class='amt'>{rp(e['jumlah'])}</div>{sisa_h}</div></div>"
            f"</div></li>")


def render(pinjaman, events, warn, today, ring, cfg):
    total, bayar, sisa = ring["total"], ring["bayar"], ring["sisa"]
    persen = (bayar / total * 100) if total else 0
    awal = min((e["tgl"] for e in events if e["jenis"] == "pinjam"), default=today)
    judul = escape(cfg["judul"])
    nama_list = list(pinjaman)

    # ---- struktur kolom: [sub-cabang..., utama] per pemberi
    kolom, grp = [], {}   # kolom: dict(i, j (None=utama), w, warna)
    data = {}
    for i, nama in enumerate(nama_list):
        semua = [x for x in events if x["pemberi"] == nama and x["jenis"] != "lunas"]
        asal = {}
        for x in semua:
            if x.get("asal"):
                asal.setdefault(x["asal"], []).append(x)
        data[nama] = (semua, asal)
        a0 = len(kolom)
        for j, nm in enumerate(asal):
            kolom.append(dict(i=i, j=j, w=W_SUB, warna=SUBWARNA[j % len(SUBWARNA)], nama=nm))
        kolom.append(dict(i=i, j=None, w=W_MAIN, warna=WARNA[i % len(WARNA)], nama=nama))
        grp[i] = (a0, len(kolom) - 1)
    ncol = len(kolom)
    W = [k["w"] for k in kolom]
    cols = " ".join(f"{x:.2f}fr" for x in W)
    S = GUT + G
    (p1, o1), (pn, on) = centro(W, *grp[0], S), centro(W, *grp[len(nama_list) - 1], S)
    mid = f"calc({(p1 + pn) / 2:.3f}% + {(o1 + on) / 2:.2f}px)"

    # ---- item per kolom: list (tgl, html)
    main_items = {i: [] for i in range(len(nama_list))}
    sub_items = {}
    fork_date = {}
    for i, nama in enumerate(nama_list):
        semua, asal = data[nama]
        done = False
        for e in semua:
            if e.get("asal"):
                if e["jenis"] == "bayar":
                    sm = e.get("sisa_main", 0)
                    main_items[i].append((e["tgl"], kartu(dict(
                        tgl=e["tgl"], jenis="bayar", jumlah=e["jumlah"],
                        ket=f"Pelunasan ke {e['asal']}" if e.get("sisa", 1) <= 0.5 else f"Cicilan ke {e['asal']}",
                        src=f"Hutang Bunda berkurang: {rp(sm + e['jumlah'])} → {rp(sm)}", sisa=sm)), ""))
                if done:
                    continue
                done = True
                pj_ev = [x for x in semua if x.get("asal") and x["jenis"] == "pinjam"]
                ket0 = pj_ev[0]["ket"] if pj_ev and pj_ev[0]["ket"] else "Pinjaman"
                main_items[i].append((e["tgl"], kartu(dict(
                    tgl=e["tgl"], jenis="pinjam", jumlah=sum(x["jumlah"] for x in pj_ev),
                    ket=f"{ket0} (dari {len(asal)} pihak)", sisa=pj_ev[-1]["sisa_main"] if pj_ev else 0)), "fork"))
                fork_date[i] = e["tgl"]
            else:
                main_items[i].append((e["tgl"], kartu(e), ""))
        for j, (nm, evs) in enumerate(asal.items()):
            sub_items[(i, j)] = [(x["tgl"], kartu(x)) for x in evs]

    # ---- baris
    def kol_idx(i, j):
        a0, b0 = grp[i]
        return b0 if j is None else a0 + j

    dates = sorted({d for v in main_items.values() for d, *_ in v} | {d for v in sub_items.values() for d, _ in v})
    rows = []  # dict(label, cells{col:[html]}, kind, fork_i)
    for d in dates:
        forks = [i for i in fork_date if fork_date[i] == d]
        A = dict(label=d, cells={}, kind="n")
        C = dict(label=None, cells={}, kind="n")
        for i in main_items:
            its = [x for x in main_items[i] if x[0] == d]
            if i in forks:
                k = next(n for n, x in enumerate(its) if x[2] == "fork")
                A["cells"].setdefault(kol_idx(i, None), []).extend(x[1] for x in its[:k])
                C["cells"].setdefault(kol_idx(i, None), []).extend(x[1] for x in its[k:])
            elif its:
                A["cells"].setdefault(kol_idx(i, None), []).extend(x[1] for x in its)
        for (i, j), v in sub_items.items():
            its = [x[1] for x in v if x[0] == d]
            if its:
                tgt = C if i in forks else A
                tgt["cells"].setdefault(kol_idx(i, j), []).extend(its)
        rows.append(A)
        if forks:
            rows.append(dict(label=None, cells={}, kind="fork", fork_i=forks))
            rows.append(C)
    # penanda hari ini setelah baris-baris bertanggal <= hari ini
    pos = len(rows)
    for n, r in enumerate(rows):
        if r["label"] and r["label"] > today:
            pos = n
            break
    rows.insert(pos, dict(label=today, cells={}, kind="today"))
    # rel: baris fork per pemberi & baris terakhir tiap sub
    fork_row = {}
    for n, r in enumerate(rows):
        if r["kind"] == "fork":
            for i in r["fork_i"]:
                fork_row[i] = n
    last_row = {}
    for n, r in enumerate(rows):
        for c in r["cells"]:
            last_row[c] = n

    # ---- HTML
    h = [f"<h1>{judul}</h1><div class='mu'>Diperbarui {tgl(today)}</div>"]
    h.append("<div class='sum'>"
             f"<div class='s-pinjam'><span class='mu'>Total pinjaman</span><b>{rp(total)}</b></div>"
             f"<div class='s-bayar'><span class='mu'>Sudah dibayar</span><b>{rp(bayar)}</b></div>"
             f"<div class='s-sisa'><span class='mu'>Sisa hutang</span><b>{rp(sisa)}</b></div></div>")
    h.append(f"<div class='bar'><i style='width:{persen:.1f}%'></i></div>"
             f"<div class='mu' style='margin-top:4px'>{persen:.1f}% terbayar</div>")
    for x in warn:
        h.append(f"<div class='warn'>{escape(x)}</div>")
    h.append(blok_pembelian(cfg.get('beli', []), cfg['debitur']))
    h.append(f"<div class='trunk' style='--m:{mid}'><div class='tr'><span>{escape(cfg['debitur'])} · Hutang awal · {tgl(awal)}</span>"
             f"<b>{rp(total)}</b></div></div><div class='stem' style='--m:{mid}'></div>")
    h.append(f"<div class='tgrid' style='--cols:{cols}'>")

    n = len(nama_list)
    for i, nama in enumerate(nama_list):
        pj = pinjaman[nama]
        sisa_p = pj["pokok"] - pj["dibayar"]
        pr = (pj["dibayar"] / pj["pokok"] * 100) if pj["pokok"] else 0
        info = ("<span class='badge'>✓ LUNAS</span>" if sisa_p <= 0.5
                else f"<span class='bi'>sisa {rp(sisa_p)} dari {rp(pj['pokok'])}</span>")
        kiri = "50%" if i == 0 else f"-{G // 2}px"
        kanan = "50%" if i == n - 1 else f"-{G // 2}px"
        garis = "" if n == 1 else f"<div class='bl' style='left:{kiri};right:{kanan}'></div>"
        a0, b0 = grp[i]
        h.append(f"<div class='branch' style='--bc:{WARNA[i % len(WARNA)]};grid-column:{a0 + 2}/{b0 + 3};grid-row:1;order:{i * 100000}'>"
                 f"{garis}<div class='bs'></div>"
                 f"<div class='bh'><div class='r'><span class='nm'>{escape(nama)}</span>{info}</div>"
                 f"<div class='bbar'><i style='width:{pr:.1f}%'></i></div></div></div>")

    for r, row in enumerate(rows):
        if row["label"] is not None:
            h.append(f"<div class='dg{' dnow' if row['kind'] == 'today' else ''}' "
                     f"style='grid-column:1;grid-row:{r + 2}'>{tgl(row['label'])}</div>")
        for c, k in enumerate(kolom):
            i, j = k["i"], k["j"]
            if j is None:
                order = i * 100000 + (r + 1) * 100
                rail = True
            else:
                fr = fork_row.get(i, 0)
                order = i * 100000 + (fr + 1) * 100 + 10 + j * 20 + min(max(r - fr, 0), 19)
                rail = fr < r <= last_row.get((c), -1)
            pos_css = f"grid-column:{c + 2};grid-row:{r + 2};order:{order}"
            if row["kind"] == "fork" and i in row["fork_i"]:
                if j is not None:
                    asal = data[nama_list[i]][1]
                    nm = list(asal)[j]
                    evs = asal[nm]
                    pok = sum(x["jumlah"] for x in evs if x["jenis"] == "pinjam")
                    byr = sum(x["jumlah"] for x in evs if x["jenis"] == "bayar")
                    sa = pok - byr
                    info = ("<span class='badge'>✓ LUNAS</span>" if sa <= 0.5
                            else f"<span class='bi'>sisa {rp(sa)} dari {rp(pok)}</span>")
                    ksub = max(jj for jj in range(len(asal)))
                    h.append(f"<div class='sfc' style='--sc:{k['warna']};{pos_css}'>"
                             f"<div class='fl' style='left:{'50%' if j == 0 else '-14px'};right:-14px'></div>"
                             f"<div class='bs2'></div><div class='sh'><span class='nm'>{escape(nm)}</span>{info}</div></div>")
                    continue
                h.append(f"<ul class='tl cell' style='--bc:{k['warna']};{pos_css}'><li class='fseg'></li></ul>")
                continue
            if row["kind"] == "today":
                isi = "<li class='now'>Hari ini</li>" if j is None else ""
            else:
                isi = "".join(row["cells"].get(c, []))
            cls = "tl cell" + ("" if rail else " norail") + ("" if isi else " emp")
            h.append(f"<ul class='{cls}' style='--bc:{k['warna']};{pos_css}'>{isi}</ul>")
    h.append("</div>")

    return (f"<!doctype html><html lang='id'><head><meta charset='utf-8'>"
            f"<meta name='color-scheme' content='light'>"
            f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
            f"<title>{judul}</title><style>{CSS}</style></head>"
            f"<body><main>{''.join(h)}</main></body></html>")


def baca_pengaturan(path):
    cfg = {"judul": "Timeline hutang rumah", "debitur": "Penanggung hutang"}
    wb = load_workbook(path, data_only=True)
    if "Pengaturan" in wb.sheetnames:
        for r in wb["Pengaturan"].iter_rows(min_row=2, values_only=True):
            if r and r[0] and r[1]:
                k = str(r[0]).strip().lower()
                if k == "judul halaman":
                    cfg["judul"] = str(r[1]).strip()
                elif k == "penanggung hutang":
                    cfg["debitur"] = str(r[1]).strip()
    return cfg


def baca_pembelian(path):
    wb = load_workbook(path, data_only=True)
    if "Pembelian" not in wb.sheetnames:
        return []
    out = []
    for r in wb["Pembelian"].iter_rows(min_row=2, max_col=5, values_only=True):
        if r and r[0] and r[1] and isinstance(r[3], (int, float)):
            out.append(dict(grup=str(r[0]).strip(), item=str(r[1]).strip(), sumber=str(r[2] or "").strip(),
                            jumlah=float(r[3]), rinci=str(r[4]).strip() if r[4] else ""))
    return out


def blok_pembelian(rows, debitur):
    if not rows:
        return ""
    tot = sum(x["jumlah"] for x in rows)
    pj = sum(x["jumlah"] for x in rows if x["sumber"] != debitur)
    sendiri = tot - pj
    grup = {}
    for x in rows:
        grup.setdefault(x["grup"], []).append(x)
    h = [f"<section class='beli'><div class='bh2'><span>Rincian pembelian rumah</span><b>{rp(tot)}</b></div>",
         f"<div class='sbar'><i class='b1' style='width:{pj / tot * 100:.1f}%'></i><i class='b2' style='width:{sendiri / tot * 100:.1f}%'></i></div>",
         f"<div class='leg'><span><i class='b1'></i>Dari pinjaman {rp(pj)}</span><span><i class='b2'></i>Dana sendiri {rp(sendiri)}</span></div>",
         "<div class='bg'>"]
    for g_, its in grup.items():
        h.append(f"<div class='bcol'><div class='bt'><span>{escape(g_)}</span><b>{rp(sum(x['jumlah'] for x in its))}</b></div>")
        for x in its:
            cls = "p" if x["sumber"] != debitur else "s"
            rinci = f"<div class='ri'>{escape(x['rinci'])}</div>" if x["rinci"] else ""
            h.append(f"<div class='bi2 {cls}'><div><div>{escape(x['item'])}</div><div class='ri'>{escape(x['sumber'])}</div>{rinci}</div><b>{rp(x['jumlah'])}</b></div>")
        h.append("</div>")
    h.append("</div></section>")
    return "".join(h)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("excel", nargs="?", default="data_hutang.xlsx")
    ap.add_argument("-o", "--output", default="index.html")
    ap.add_argument("--today", help="YYYY-MM-DD (untuk uji coba)")
    a = ap.parse_args()
    today = datetime.strptime(a.today, "%Y-%m-%d").date() if a.today else date.today()

    pinjaman, events, warn, ring = muat(a.excel, today)
    Path(a.output).write_text(render(pinjaman, events, warn, today, ring, dict(baca_pengaturan(a.excel), beli=baca_pembelian(a.excel))), encoding="utf-8")
    for w in warn:
        print("PERINGATAN:", w)
    print(f"OK: {a.output} ({len(events)} kejadian)")


if __name__ == "__main__":
    main()
