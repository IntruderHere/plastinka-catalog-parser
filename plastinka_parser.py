#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер полного каталога винила Plastinka.com  (https://plastinka.com/lp)
=======================================================================

Запуск (ничего заранее устанавливать не нужно):

    python plastinka_parser.py

Скрипт сам:
  1) проверит и при необходимости поставит зависимости (httpx, lxml, openpyxl);
  2) возьмёт список всех карточек винила из https://plastinka.com/sitemap.xml;
  3) обойдёт карточки с вежливой задержкой, ретраями и автопаузой при 403/429;
  4) сохранит результат в  catalog.jsonl / Каталог_винила_Plastinka_<дата>.csv/.xlsx.

Полезные ключи:
    --limit 50           только первые 50 карточек (проверка)
    --workers 3          параллельных потоков (по умолчанию 3)
    --delay 0.5          пауза потока между запросами, сек (по умолчанию 0.5)
    --out DIR            папка для результатов (по умолчанию текущая)
    --max-minutes 30     остановиться через N минут и собрать файлы
    --sitemap FILE/URL   свой sitemap (по умолчанию скачивается с сайта)

Прервать (Ctrl+C) безопасно: уже собранное сохраняется, файлы собираются.
Повторный запуск продолжает с места остановки (дедупликация по id карточки).

Данные по каждой пластинке: исполнитель, альбом, год, страна, лейбл, стиль,
тип (оригинал/переиздание), состояние, размер, ЦЕНА, старая цена, скидка в рублях
и процентах, наличие, примечания, ссылка на карточку, ссылка на обложку.

robots.txt сайта: карточки /lp/item/... и sitemap разрешены; пагинация листинга
(?page=) запрещена — поэтому обход идёт по sitemap, а не по страницам листинга.
"""
from __future__ import annotations

import argparse
import asyncio
import json
import random
import re
import subprocess
import sys
import time
from datetime import date
from pathlib import Path

# ----------------------------------------------------------------------------
# 0. Автоустановка зависимостей
# ----------------------------------------------------------------------------
REQUIRED = [("httpx", "httpx"), ("lxml", "lxml"),
            ("cssselect", "cssselect"), ("openpyxl", "openpyxl")]


def _try_import(name: str) -> bool:
    try:
        __import__(name)
        return True
    except ImportError:
        return False


def ensure_dependencies() -> None:
    missing = [pkg for mod, pkg in REQUIRED if not _try_import(mod)]
    if not missing:
        return
    print(f"[deps] не хватает: {', '.join(missing)} — устанавливаю…", flush=True)

    # pip может отсутствовать в свежем venv
    try:
        subprocess.run([sys.executable, "-m", "pip", "--version"],
                       check=True, capture_output=True)
    except Exception:
        subprocess.run([sys.executable, "-m", "ensurepip", "--upgrade"], check=False)

    variants = [
        [sys.executable, "-m", "pip", "install", "--quiet", *missing],
        [sys.executable, "-m", "pip", "install", "--quiet", "--user", *missing],
        [sys.executable, "-m", "pip", "install", "--quiet",
         "--break-system-packages", *missing],
    ]
    for cmd in variants:
        if subprocess.call(cmd) == 0:
            break
    else:
        sys.exit("[deps] не удалось установить зависимости. Установите вручную:\n"
                 f"        {sys.executable} -m pip install {' '.join(missing)}")

    import importlib
    importlib.invalidate_caches()
    still = [pkg for mod, pkg in REQUIRED if not _try_import(mod)]
    if still:
        sys.exit(f"[deps] всё ещё нет модулей: {still}. Запустите вручную pip install {still}")
    print("[deps] готово", flush=True)


ensure_dependencies()

import httpx            # noqa: E402
import lxml.html        # noqa: E402
from openpyxl import Workbook                                        # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter                         # noqa: E402

# ----------------------------------------------------------------------------
# 1. Настройки
# ----------------------------------------------------------------------------
SITE = "https://plastinka.com"
SITEMAP_URL = f"{SITE}/sitemap.xml"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36")
TODAY = date.today().strftime("%Y-%m-%d")

LABELS = {
    "исполнитель": "artist",
    "альбом": "album",
    "страна": "country",
    "лейбл": "label",
    "год выпуска/год издания": "year",
    "год выпуска": "year",
    "стиль": "style",
    "тип": "edition",
    "состояние пластинки/конверта": "condition",
    "размер диска": "size",
}
COND_RE = re.compile(r"([A-Z]{1,3}[+-]?/[A-Z]{1,3}[+-]?)(?![A-Z/])")


# ----------------------------------------------------------------------------
# 2. Sitemap
# ----------------------------------------------------------------------------
async def fetch_sitemap(client: httpx.AsyncClient, src: str) -> list[str]:
    if src.startswith("http"):
        r = await client.get(src, timeout=60)
        r.raise_for_status()
        xml = r.text
    else:
        xml = Path(src).read_text(encoding="utf-8", errors="ignore")
    urls = re.findall(r"<loc>(https://plastinka\.com/lp/item/[^<]+)</loc>", xml)
    seen, out = set(), []
    for u in urls:
        if u not in seen:
            seen.add(u)
            out.append(u)
    return out


# ----------------------------------------------------------------------------
# 3. Парсер карточки
# ----------------------------------------------------------------------------
def parse_item(page: str, url: str) -> dict:
    doc = lxml.html.fromstring(page)
    for bad in doc.cssselect("span.tooltip-html, i.icon-info, svg"):
        try:
            bad.getparent().remove(bad)
        except Exception:
            pass

    m_id = re.search(r"/item/(\d+)", url)
    rec: dict = {"url": url, "id": m_id.group(1) if m_id else None}
    h1 = doc.cssselect("h1.first-title")
    rec["name"] = " ".join(h1[0].text_content().split()) if h1 else ""

    fields: dict[str, str] = {}
    p = doc.cssselect('article.txt-product p[itemprop="description"]')
    if p:
        html = lxml.html.tostring(p[0], encoding="unicode")
        html = re.sub(r"<br\s*/?>", "\n", html)
        chunks = [re.sub(r"\s+", " ", c).strip()
                  for c in lxml.html.fromstring(html).text_content().split("\n")]
        pending = None
        for ch in chunks:
            if not ch:
                continue
            key = ch.rstrip(":").strip().lower()
            if ch.endswith(":") and key in LABELS:
                pending = LABELS[key]
                continue
            if key in LABELS and len(ch) <= 40:
                pending = LABELS[key]
                continue
            if pending:
                fields.setdefault(pending, ch)
                pending = None
        for ch in chunks:                      # «Лейбл: Arista» в одной строке
            if ":" in ch:
                k, v = ch.split(":", 1)
                k = k.strip().lower()
                if k in LABELS and v.strip():
                    fields.setdefault(LABELS[k], v.strip())

    for k in ("artist", "album", "country", "label", "year", "style",
              "edition", "condition", "size"):
        rec[k] = fields.get(k, "")

    if rec["condition"]:
        m = COND_RE.search(rec["condition"])
        rec["condition"] = m.group(1) if m else rec["condition"][:12]
    if rec["year"]:
        m = re.search(r"(19|20)\d{2}", rec["year"])
        rec["year"] = m.group(0) if m else ""
    if not rec["artist"] and "—" in rec["name"]:
        a, _, b = rec["name"].partition("—")
        rec["artist"], rec["album"] = a.strip(), b.strip()

    # ---- цены ----
    box = doc.cssselect("div.product-card__price")
    scope = box[0] if box else doc
    meta = scope.cssselect('meta[itemprop="price"]')
    if meta and (meta[0].get("content") or "").isdigit():
        rec["price_rub"] = int(meta[0].get("content"))
    else:
        rec["price_rub"] = None
    old = scope.cssselect("span.prev-price")
    rec["old_price_rub"] = int(re.sub(r"\D", "", old[0].text_content())) if old else None
    if rec["price_rub"] and rec["old_price_rub"]:
        rec["discount_rub"] = rec["old_price_rub"] - rec["price_rub"]
        rec["discount_pct"] = round(100 * rec["discount_rub"] / rec["old_price_rub"])
    else:
        rec["discount_rub"] = rec["discount_pct"] = None
    cur = scope.cssselect('span[itemprop="priceCurrency"]')
    rec["currency"] = cur[0].get("content") if cur else "RUB"

    rec["in_stock"] = "OutOfStock" not in page
    rec["availability"] = ('В корзину' if 'data-action="to-cart"' in page else
                           ("Под заказ" if "Под заказ" in page else ""))
    img = doc.cssselect('link[itemprop="contentUrl"]')
    rec["image"] = img[0].get("href") if img else ""
    bq = doc.cssselect("article.txt-product blockquote")
    rec["notes"] = " ".join(bq[0].text_content().split()) if bq else ""
    rec["fetched_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    return rec


# ----------------------------------------------------------------------------
# 4. Обход
# ----------------------------------------------------------------------------
def load_done(path: Path) -> set[str]:
    done: set[str] = set()
    if path.exists():
        with path.open(encoding="utf-8") as fh:
            for line in fh:
                try:
                    done.add(json.loads(line)["id"])
                except Exception:
                    pass
    return done


async def worker(queue: asyncio.Queue, client: httpx.AsyncClient, args,
                 done: set[str], stats: dict, stop: asyncio.Event, out: Path,
                 on_item=None) -> None:
    err_path = out.with_suffix(".errors")
    while not stop.is_set():
        try:
            url = await asyncio.wait_for(queue.get(), timeout=3)
        except asyncio.TimeoutError:
            return

        m = re.search(r"/item/(\d+)", url)
        rid = m.group(1) if m else url
        if rid in done:
            queue.task_done()
            continue

        page, err, status = None, None, None
        for attempt in range(1, args.retries + 1):
            try:
                r = await client.get(url, timeout=args.timeout)
                status = r.status_code
                if status == 200:
                    page = r.text
                    break
                if status in (403, 429, 503):
                    stats["throttled"] += 1
                    await asyncio.sleep(min(60, 5 * attempt) + random.random() * 3)
                    continue
                err = f"HTTP {status}"
                break
            except Exception as exc:
                err = f"{type(exc).__name__}: {exc}"[:120]
                await asyncio.sleep(2 ** attempt + random.random())

        if page:
            try:
                rec = parse_item(page, url)
                with out.open("a", encoding="utf-8") as fh:
                    fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
                done.add(rid)
                stats["ok"] += 1
                if on_item:
                    on_item(rec)
            except Exception as exc:
                stats["parse_failed"] += 1
                print(f"[parse] {url}: {type(exc).__name__}: {exc}", file=sys.stderr)
        else:
            stats["failed"] += 1
            with err_path.open("a", encoding="utf-8") as fh:
                fh.write(f"{url}\t{status}\t{err}\n")

        n = stats["ok"] + stats["failed"]
        if n and n % 50 == 0:
            el = max(1e-6, time.time() - stats["t0"])
            rate = n / el
            left = (stats["total"] - n) / rate / 60 if rate else 0
            print(f"[{stats['ok']}/{stats['total']}] {rate:.1f} карт./с | "
                  f"ошибок {stats['failed']} | осталось ~{left:.0f} мин", flush=True)
        queue.task_done()
        await asyncio.sleep(args.delay)


# ----------------------------------------------------------------------------
# 5. Выгрузка xlsx + csv
# ----------------------------------------------------------------------------
HEAD = ["№", "Исполнитель", "Альбом", "Год", "Страна", "Лейбл", "Стиль", "Тип",
        "Состояние", "Размер", "Цена, ₽", "Было, ₽", "Скидка, ₽", "Скидка, %",
        "Наличие", "Примечания", "Ссылка"]
WIDTH = [5, 26, 46, 7, 15, 22, 24, 12, 11, 8, 11, 11, 11, 10, 14, 40, 62]


def collect_rows(jsonl: Path) -> list[dict]:
    """Читает catalog.jsonl, дедуплицирует по id (последняя запись побеждает)."""
    rows_by_id: dict[str, dict] = {}
    if jsonl.exists():
        with jsonl.open(encoding="utf-8") as fh:
            for line in fh:
                try:
                    r = json.loads(line)
                    rows_by_id[r.get("id") or r.get("url")] = r
                except Exception:
                    pass
    return list(rows_by_id.values())


def export(rows: list[dict], out_dir: Path) -> tuple[Path, Path]:
    from collections import Counter

    rows = sorted(rows, key=lambda r: (r.get("artist") or "").lower())
    wb = Workbook()
    ws = wb.active
    ws.title = "Каталог LP"
    fill = PatternFill("solid", fgColor="1F3864")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c, (name, w) in enumerate(zip(HEAD, WIDTH), 1):
        cell = ws.cell(row=1, column=c, value=name)
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 30

    for i, r in enumerate(rows, 2):
        vals = [i - 1, r.get("artist"), r.get("album"),
                int(r["year"]) if r.get("year") else None,
                r.get("country") or "—", r.get("label"), r.get("style") or "—",
                r.get("edition"), r.get("condition"), r.get("size"),
                r.get("price_rub"), r.get("old_price_rub"), r.get("discount_rub"),
                (r["discount_pct"] / 100 if r.get("discount_pct") else None),
                r.get("availability") or "—", r.get("notes"), r.get("url")]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font = Font(size=10)
            cell.border = border
            cell.alignment = Alignment(vertical="top", wrap_text=c in (2, 3, 7, 16))
            if c in (11, 12, 13):
                cell.number_format = "# ##0"
                cell.alignment = Alignment(horizontal="right", vertical="top")
            if c == 14:
                cell.number_format = "0%"
                cell.alignment = Alignment(horizontal="right", vertical="top")
            if c == 17:
                cell.hyperlink = v
                cell.font = Font(size=9, color="1155CC", underline="single")
            if c == 1:
                cell.alignment = Alignment(horizontal="center", vertical="top")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEAD))}{len(rows)+1}"

    ws2 = wb.create_sheet("Сводка")
    prices = [r["price_rub"] for r in rows if r.get("price_rub")]
    disc = [r["discount_rub"] for r in rows if r.get("discount_rub")]
    ws2["A1"] = f"Каталог винила Plastinka.com — {len(rows)} позиций ({TODAY})"
    ws2["A1"].font = Font(bold=True, size=13)
    ws2.column_dimensions["A"].width = 40
    ws2.column_dimensions["B"].width = 34
    summary = [
        ("Источник", "https://plastinka.com/lp (sitemap → карточки)"),
        ("Позиций", len(rows)),
        ("С ценой", len(prices)),
        ("Цена: мин / средняя / макс, ₽",
         f"{min(prices)} / {round(sum(prices)/len(prices))} / {max(prices)}" if prices else "—"),
        ("Сумма каталога по текущим ценам, ₽",
         f"{sum(prices):,}".replace(",", " ") if prices else "—"),
        ("Со скидкой", len(disc)),
        ("Сумма скидок, ₽", f"{sum(disc):,}".replace(",", " ") if disc else 0),
        ("Пустой год", sum(1 for r in rows if not r.get("year"))),
        ("Пустой стиль", sum(1 for r in rows if not r.get("style"))),
    ]
    for i, (k, v) in enumerate(summary, 3):
        ws2.cell(row=i, column=1, value=k).font = Font(bold=True, size=10)
        ws2.cell(row=i, column=2, value=v).font = Font(size=10)

    row0 = len(summary) + 5
    for title, counter in [("Наличие", Counter(r.get("availability") or "—" for r in rows)),
                           ("Тип", Counter(r.get("edition") or "—" for r in rows)),
                           ("Состояние", Counter(r.get("condition") or "—" for r in rows)),
                           ("Страна", Counter(r.get("country") or "не указана" for r in rows)),
                           ("Стиль", Counter(r.get("style") or "—" for r in rows))]:
        ws2.cell(row=row0, column=1, value=title).font = Font(bold=True, size=11)
        for j, (k, n) in enumerate(counter.most_common(10), 1):
            ws2.cell(row=row0 + j, column=1, value=k).font = Font(size=10)
            ws2.cell(row=row0 + j, column=2, value=n).font = Font(size=10)
        row0 += len(counter.most_common(10)) + 3

    xlsx = out_dir / f"Каталог_винила_Plastinka_{TODAY}.xlsx"
    wb.save(xlsx)

    import csv
    csv_path = out_dir / f"Каталог_винила_Plastinka_{TODAY}.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(HEAD)
        for i, r in enumerate(rows, 1):
            w.writerow([i, r.get("artist"), r.get("album"), r.get("year"),
                        r.get("country"), r.get("label"), r.get("style"),
                        r.get("edition"), r.get("condition"), r.get("size"),
                        r.get("price_rub"), r.get("old_price_rub"),
                        r.get("discount_rub"), r.get("discount_pct"),
                        r.get("availability"), r.get("notes"), r.get("url")])
    return xlsx, csv_path


# ----------------------------------------------------------------------------
# 6. main
# ----------------------------------------------------------------------------
async def run(args) -> int:
    out_dir = Path(args.out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    jsonl = out_dir / "catalog.jsonl"

    headers = {"User-Agent": UA, "Accept-Language": "ru-RU,ru;q=0.9",
               "Accept": "text/html,application/xhtml+xml",
               "Accept-Encoding": "gzip, deflate"}
    limits = httpx.Limits(max_connections=args.workers + 2)

    async with httpx.AsyncClient(follow_redirects=True, limits=limits,
                                 headers=headers, timeout=args.timeout) as client:
        print("[sitemap] читаю список карточек…", flush=True)
        urls = await fetch_sitemap(client, args.sitemap)
        if not urls:
            print("[sitemap] не нашёл ни одной карточки — проверьте доступ к сайту",
                  file=sys.stderr)
            return 2
        if args.limit:
            urls = urls[:args.limit]
        done = load_done(jsonl)
        todo = [u for u in urls if (re.search(r"/item/(\d+)", u) or [None, u])[1] not in done]
        print(f"[sitemap] всего карточек: {len(urls)} | уже собрано: {len(done)} | "
              f"к обходу: {len(todo)}", flush=True)
        if not todo:
            print("[crawl] всё уже собрано, только пересобираю файлы", flush=True)
        else:
            print(f"[crawl] {args.workers} потоков, пауза {args.delay} с — старт", flush=True)

        queue: asyncio.Queue = asyncio.Queue()
        for u in todo:
            queue.put_nowait(u)

        stats = {"ok": 0, "failed": 0, "parse_failed": 0, "throttled": 0,
                 "total": len(todo), "t0": time.time()}
        stop = asyncio.Event()

        deadline = (time.time() + args.max_minutes * 60) if args.max_minutes else None

        async def timer():
            while not stop.is_set():
                await asyncio.sleep(5)
                if deadline and time.time() > deadline:
                    print(f"\n[time] лимит {args.max_minutes} мин — останавливаюсь", flush=True)
                    stop.set()

        tasks = [asyncio.create_task(worker(queue, client, args, done, stats, stop, jsonl))
                 for _ in range(args.workers)]
        if deadline:
            tasks.append(asyncio.create_task(timer()))
        try:
            await asyncio.gather(*tasks)
        except (KeyboardInterrupt, asyncio.CancelledError):
            stop.set()
            print("\n[crawl] прервано пользователем", flush=True)

    rows = collect_rows(jsonl)
    if not rows:
        print("[export] данных нет", file=sys.stderr)
        return 1

    xlsx, csv_path = export(rows, out_dir)
    prices = [r["price_rub"] for r in rows if r.get("price_rub")]
    disc = [r["discount_rub"] for r in rows if r.get("discount_rub")]
    print("\n" + "=" * 60)
    print(f"собрано позиций: {len(rows)}  (ok={stats['ok']}, ошибок={stats['failed']})")
    print(f"с ценой: {len(prices)}/{len(rows)} | "
          + (f"цена {min(prices)}–{max(prices)} ₽, средняя "
             f"{round(sum(prices)/len(prices))} ₽" if prices else "цен нет"))
    if disc:
        print(f"со скидкой: {len(disc)} поз., сумма скидок "
              f"{sum(disc):,}".replace(",", " ") + " ₽")
    print(f"xlsx: {xlsx}")
    print(f"csv : {csv_path}")
    print("=" * 60)
    return 0


def main() -> int:
    p = argparse.ArgumentParser(
        description="Полный каталог винила Plastinka.com → xlsx/csv",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    p.add_argument("--limit", type=int, default=0, help="0 = весь каталог")
    p.add_argument("--workers", type=int, default=3)
    p.add_argument("--delay", type=float, default=0.5, help="пауза потока между запросами, с")
    p.add_argument("--timeout", type=float, default=40.0)
    p.add_argument("--retries", type=int, default=3)
    p.add_argument("--max-minutes", type=float, default=0, help="0 = без лимита")
    p.add_argument("--out", default=".", help="папка для results")
    p.add_argument("--sitemap", default=SITEMAP_URL)
    args = p.parse_args()
    try:
        return asyncio.run(run(args))
    except KeyboardInterrupt:
        print("\n[crawl] прервано — собираю файлы из уже собранного…", flush=True)
        out_dir = Path(args.out).expanduser().resolve()
        rows = collect_rows(out_dir / "catalog.jsonl")
        if rows:
            xlsx, csv_path = export(rows, out_dir)
            print(f"сохранено позиций: {len(rows)}\n  xlsx: {xlsx}\n  csv : {csv_path}")
            print("Повторный запуск продолжит с места остановки.")
        else:
            print("данных ещё нет")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
