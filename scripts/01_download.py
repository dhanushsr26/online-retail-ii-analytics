"""
Step 1a - Acquire the raw dataset.

Downloads UCI Online Retail II (dataset 502) and extracts it to 00_source/.
Idempotent: skips the download if the .xlsx is already present.
"""
import hashlib
import pathlib
import urllib.request
import zipfile

URL = "https://archive.ics.uci.edu/static/public/502/online+retail+ii.zip"
ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "00_source"
ZIP_PATH = SRC / "online_retail_II.zip"
XLSX_PATH = SRC / "online_retail_II.xlsx"


def main() -> None:
    SRC.mkdir(parents=True, exist_ok=True)

    if XLSX_PATH.exists():
        print(f"Already present: {XLSX_PATH.name} "
              f"({XLSX_PATH.stat().st_size / 1024 / 1024:.1f} MB) - skipping download")
        return

    print(f"Downloading {URL}")
    req = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=180) as resp, open(ZIP_PATH, "wb") as fh:
        fh.write(resp.read())

    size_mb = ZIP_PATH.stat().st_size / 1024 / 1024
    sha = hashlib.sha256(ZIP_PATH.read_bytes()).hexdigest()
    print(f"Downloaded {ZIP_PATH.name}  {size_mb:.1f} MB")
    print(f"SHA256 {sha}")

    with zipfile.ZipFile(ZIP_PATH) as zf:
        print("Archive contents:")
        for info in zf.infolist():
            print(f"  {info.filename}  {info.file_size / 1024 / 1024:.1f} MB")
        zf.extractall(SRC)

    if not XLSX_PATH.exists():
        # Archive may name the file differently; normalise it.
        candidates = [p for p in SRC.glob("*.xlsx")]
        if len(candidates) == 1:
            candidates[0].rename(XLSX_PATH)
        else:
            raise SystemExit(f"Expected one .xlsx after extract, found: {candidates}")

    print(f"Ready: {XLSX_PATH}  ({XLSX_PATH.stat().st_size / 1024 / 1024:.1f} MB)")


if __name__ == "__main__":
    main()
