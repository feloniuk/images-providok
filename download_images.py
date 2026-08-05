#!/usr/bin/env python3
import os
import re
import sys
import time
import openpyxl
import urllib.request
import urllib.error
from urllib.parse import urlsplit

XLSX_PATH = "e362b4e981110dc6a83c30d165c014c1 (1).xlsx"
OUT_DIR = "photos"
LOG_PATH = "download_log.csv"

os.makedirs(OUT_DIR, exist_ok=True)

FS_UNSAFE = re.compile(r'[\\/:*?"<>|]')


def sanitize(article: str) -> str:
    return FS_UNSAFE.sub("_", article.strip())


def ext_from_url(url: str) -> str:
    path = urlsplit(url).path
    ext = os.path.splitext(path)[1]
    if not ext or len(ext) > 5:
        ext = ".jpg"
    return ext.lower()


def download(url: str, dest_path: str, retries: int = 3) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    for attempt in range(1, retries + 1):
        try:
            with urllib.request.urlopen(req, timeout=30) as resp, open(dest_path, "wb") as f:
                f.write(resp.read())
            return "OK"
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError) as e:
            if attempt == retries:
                return f"FAIL: {e}"
            time.sleep(2)


def main():
    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True)
    ws = wb["Sheet1"]

    rows = list(ws.iter_rows(values_only=True))
    header, data_rows = rows[0], rows[1:]

    log_lines = ["article,type,index,url,filename,status"]
    total = len(data_rows)

    for i, row in enumerate(data_rows, start=1):
        article_raw = row[0]
        photo = row[15]
        gallery = row[16]

        if not article_raw:
            continue
        article = sanitize(str(article_raw))

        photo_urls = [u.strip() for u in re.split(r"[;\n]+", photo)] if photo else []
        gallery_urls = [u.strip() for u in re.split(r"[;\n]+", gallery)] if gallery else []
        all_urls = [u for u in (photo_urls + gallery_urls) if u]

        if not all_urls:
            continue

        # First url overall = main photo (no suffix); rest = gallery (@1, @2, ...)
        main_url, *rest_urls = all_urls

        ext = ext_from_url(main_url)
        fname = f"{article}{ext}"
        dest = os.path.join(OUT_DIR, fname)
        status = download(main_url, dest)
        log_lines.append(f'"{article}",main,,"{main_url}","{fname}",{status}')
        print(f"[{i}/{total}] {fname}: {status}")

        for n, gurl in enumerate(rest_urls, start=1):
            ext = ext_from_url(gurl)
            fname = f"{article}@{n}{ext}"
            dest = os.path.join(OUT_DIR, fname)
            status = download(gurl, dest)
            log_lines.append(f'"{article}",gallery,{n},"{gurl}","{fname}",{status}')
            print(f"[{i}/{total}] {fname}: {status}")

    with open(LOG_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines) + "\n")

    print("Done. Log written to", LOG_PATH)


if __name__ == "__main__":
    main()
