"""Excel -> data.json çevirici.

Kullanım:
    pip install openpyxl
    python convert.py malzemeler.xlsx [data.json]

Excel'in ilk satırı başlık olmalı:
Malzeme No | Malzeme Açıklaması | Eldeki Miktar | Lokasyon No
(Başka sütunlar varsa yok sayılır.)

Aynı malzeme farklı lokasyonlarda / farklı satırlarda olabilir:
- Malzemenin toplam miktarı tüm satırların toplamıdır.
- Aynı lokasyondaki satırlar toplanır, her lokasyon ayrı listelenir.
Çıktıya güncelleme tarihi/saati (Türkiye saati) ve Excel'deki kayıt sayısı da yazılır.
"""
import json
import sys
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook

COLS = ["Malzeme No", "Malzeme Açıklaması", "Eldeki Miktar", "Lokasyon No"]


def temiz(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def sayi(v):
    """Miktarı sayıya çevirir; çevrilemezse None döner."""
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", "."))
    except ValueError:
        return None


def yaz(x):
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    return f"{x:.3f}".rstrip("0").rstrip(".")


def main(dosya, cikti="data.json"):
    ws = load_workbook(dosya, data_only=True, read_only=True).active
    satirlar = ws.iter_rows(values_only=True)
    baslik = [temiz(h) for h in next(satirlar)]
    eksik = [c for c in COLS if c not in baslik]
    if eksik:
        sys.exit("Excel'de bu sütunlar bulunamadı: " + ", ".join(eksik))
    idx = {c: baslik.index(c) for c in COLS}

    malzemeler = {}
    kayit = hatali = 0
    for s in satirlar:
        kod = temiz(s[idx["Malzeme No"]])
        if not kod:
            continue
        kayit += 1
        m = malzemeler.setdefault(kod, {"aciklama": "", "lok": {}})
        if not m["aciklama"]:
            m["aciklama"] = temiz(s[idx["Malzeme Açıklaması"]])
        miktar = sayi(s[idx["Eldeki Miktar"]])
        if miktar is None:
            hatali += 1
            miktar = 0.0
        lok = temiz(s[idx["Lokasyon No"]]) or "-"
        m["lok"][lok] = m["lok"].get(lok, 0.0) + miktar

    veri = {}
    for kod, m in malzemeler.items():
        toplam = sum(m["lok"].values())
        veri[kod] = {
            "aciklama": m["aciklama"],
            "toplam": yaz(toplam),
            "lok": [[l, yaz(q)] for l, q in sorted(m["lok"].items())],
        }

    simdi = datetime.now(timezone(timedelta(hours=3)))  # Türkiye saati (UTC+3)
    paket = {"guncelleme": simdi.strftime("%d.%m.%Y %H:%M"), "kayit": kayit, "veri": veri}
    with open(cikti, "w", encoding="utf-8") as f:
        json.dump(paket, f, ensure_ascii=False, separators=(",", ":"))
    satir_lok = sum(len(v["lok"]) for v in veri.values())
    print(f"{kayit} kayıt okundu: {len(veri)} farklı malzeme, {satir_lok} malzeme-lokasyon çifti ({paket['guncelleme']}).")
    if hatali:
        print(f"UYARI: {hatali} satırdaki miktar sayıya çevrilemedi, 0 kabul edildi.")


if __name__ == "__main__":
    if len(sys.argv) not in (2, 3):
        sys.exit("Kullanım: python convert.py malzemeler.xlsx [data.json]")
    main(*sys.argv[1:])
