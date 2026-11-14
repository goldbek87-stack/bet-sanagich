"""
BET SANAGICH - papkadagi PDF va Word fayllarning sahifasini ochmasdan sanaydi.

Ishlatish:
  1) Shu faylni (bet_sanagich.py) mijoz fayllari turgan papkaga tashlang.
  2) Ikki marta bosing (yoki: python bet_sanagich.py).
  Boshqa papka uchun: python bet_sanagich.py "C:\\papka\\yo'li"

Kerak: pip install pypdf
Word (.docx/.doc) aniq sanalishi uchun LibreOffice o'rnatilgan bo'lishi kerak
(bepul: libreoffice.org). U bo'lmasa Word fayllar taxminiy (fayl ichidagi) son
bilan sanaladi va "~" belgisi qo'yiladi.
"""
import os
import sys
import shutil
import subprocess
import tempfile
import zipfile
import re
from concurrent.futures import ThreadPoolExecutor

try:
    from pypdf import PdfReader
except ImportError:
    print("pypdf yo'q. Buyruq oynasida yozing:  pip install pypdf")
    input("Enter bosing...")
    sys.exit(1)

IMG = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}
WORD = {".docx", ".doc", ".rtf", ".odt"}


def pdf_pages(path):
    try:
        return len(PdfReader(path).pages)
    except Exception:
        return None


def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


def install_libreoffice():
    """Windows'da winget orqali LibreOffice'ni o'zi o'rnatadi."""
    if os.name != "nt" or not shutil.which("winget"):
        print("  Avtomatik o'rnatib bo'lmadi. libreoffice.org dan o'rnating.")
        return None
    print("  LibreOffice o'rnatilmoqda (bir marta, 2-5 daqiqa). Ruxsat so'rasa 'Ha' bosing...")
    subprocess.run(["winget", "install", "--id", "TheDocumentFoundation.LibreOffice",
                    "-e", "--silent", "--accept-package-agreements",
                    "--accept-source-agreements"])
    return find_soffice()


def docx_meta_pages(path):
    """Taxminiy: Word oxirgi saqlaganda yozgan son."""
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("docProps/app.xml").decode("utf8", "ignore")
        m = re.search(r"<Pages>(\d+)</Pages>", xml)
        return int(m.group(1)) if m else None
    except Exception:
        return None


def convert_words(paths, soffice):
    """Hamma Word fayllarni bitta jarayonda PDF'ga o'tkazib sanaydi."""
    result = {}
    out = tempfile.mkdtemp(prefix="bet_")
    profile = tempfile.mkdtemp(prefix="lo_")
    try:
        uri = "file:///" + profile.replace("\\", "/").lstrip("/")
        cmd = [soffice, f"-env:UserInstallation={uri}", "--headless",
               "--convert-to", "pdf", "--outdir", out] + paths
        subprocess.run(cmd, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL, timeout=600)
        for p in paths:
            pdf = os.path.join(out, os.path.splitext(os.path.basename(p))[0] + ".pdf")
            result[p] = pdf_pages(pdf) if os.path.exists(pdf) else None
    finally:
        shutil.rmtree(out, ignore_errors=True)
        shutil.rmtree(profile, ignore_errors=True)
    return result


def main():
    if len(sys.argv) > 1:
        folder = sys.argv[1]
    elif getattr(sys, "frozen", False):
        folder = os.path.dirname(sys.executable)
    else:
        folder = os.path.dirname(os.path.abspath(__file__))
    me = os.path.basename(sys.executable if getattr(sys, "frozen", False) else __file__)
    files = sorted(f for f in os.listdir(folder)
                   if os.path.isfile(os.path.join(folder, f))
                   and not f.startswith(("~$", "."))
                   and f != me)
    pdfs = [f for f in files if f.lower().endswith(".pdf")]
    words = [f for f in files if os.path.splitext(f)[1].lower() in WORD]
    imgs = [f for f in files if os.path.splitext(f)[1].lower() in IMG]

    rows = []  # (nom, son, taxminiy_mi)
    with ThreadPoolExecutor(8) as ex:
        for f, n in zip(pdfs, ex.map(lambda x: pdf_pages(os.path.join(folder, x)), pdfs)):
            rows.append((f, n, False))

    if words:
        soffice = find_soffice()
        if not soffice:
            print("\nWord fayllarni aniq sanash uchun LibreOffice kerak, u topilmadi.")
            if input("O'zi o'rnatsinmi? (Enter = ha, n = yo'q): ").strip().lower() != "n":
                soffice = install_libreoffice()
        full = [os.path.join(folder, f) for f in words]
        exact = convert_words(full, soffice) if soffice else {}
        for f, p in zip(words, full):
            n = exact.get(p)
            if n is not None:
                rows.append((f, n, False))
            else:
                m = docx_meta_pages(p) if p.lower().endswith(".docx") else None
                rows.append((f, m, True))

    for f in imgs:
        rows.append((f, 1, False))

    print("\n" + "=" * 60)
    total, approx, bad = 0, False, []
    for f, n, ap in sorted(rows):
        if n is None:
            bad.append(f)
            print(f"  {f[:45]:<45} ???")
        else:
            total += n
            approx = approx or ap
            print(f"  {f[:45]:<45} {n:>4}{' ~' if ap else ''}")
    print("=" * 60)
    print(f"  FAYLLAR: {len(rows)} ta      JAMI: {total} bet{' (taxminiy ~)' if approx else ''}")
    if bad:
        print("  Sanab bo'lmadi:", ", ".join(bad))
    if approx:
        print("  ~ = LibreOffice yo'q yoki fayl o'qilmadi, son taxminiy.")
    print()
    if sys.stdin.isatty() or os.name == "nt":
        input("Chiqish uchun Enter bosing...")


if __name__ == "__main__":
    main()
