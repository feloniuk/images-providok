#!/usr/bin/env python3
"""
Зводить карбонові позиції 4 джерел в одну таблицю: на кожну позицію —
ціна Leti, ціна FlyShop, ціна carbon.net.ua, ціна com-posit.com.ua.

  Leti      (наш, leti.com.ua)        3ea2dc4...xlsx
  FlyShop   (наш, flyshop.com.ua)     42cebda...(1).xlsx
  carbon.net.ua    (конкурент)        carbon_net_ua_catalog.xlsx
  com-posit.com.ua (конкурент)        com_posit_carbon_catalog.xlsx

ПРИНЦИП: у таблицю потрапляють тільки значення, які буквально є у файлах-
джерелах. Нічого не домислюється і не перераховується за курсом.

Зіставлення — за "тип виробу + розміри в мм + технологія", бо артикули в
усіх чотирьох джерелах різні. Розміри беруться з колонки характеристик там,
де вона є (у конкурентів), інакше — з назви. Технологія (пултрузія / тканий
3K) читається з назви, опису й характеристик; за однакового розміру це різні
товари з різницею в ціні у 2-3 рази, тому вона входить у ключ.

Результат: price_comparison_carbon.xlsx
Запуск:    python3 build_price_comparison.py
"""
import re
import unicodedata
from collections import defaultdict

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

LETI_FILE = "3ea2dc4ac739ec2644226b46b129fa00.xlsx"
FLY_FILE = "42cebda50ca993978167698344a6f178 (1).xlsx"
CN_FILE = "carbon_net_ua_catalog.xlsx"
CP_FILE = "com_posit_carbon_catalog.xlsx"
OUT_FILE = "price_comparison_carbon.xlsx"

SRC = ["LETI", "FLY", "CN", "CP"]
SRC_TITLE = {"LETI": "Leti", "FLY": "FlyShop",
             "CN": "carbon.net.ua", "CP": "com-posit.com.ua"}

DASHES = dict.fromkeys(map(ord, "‐‑‒–—―"), "-")
DIAM = re.compile(r"[∅ØøΦφ]")
NUM = re.compile(r"\d+(?:[.,]\d+)?")


def norm_text(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s)).translate(DASHES)
    return re.sub(r"\s+", " ", s.replace("\xa0", " ")).strip()


def to_num(tok):
    if tok is None:
        return None
    m = re.search(r"-?\d+(?:\.\d+)?", norm_text(tok).replace(" ", "").replace(",", "."))
    return float(m.group()) if m else None


def nums_in(s):
    return [float(x.replace(",", ".")) for x in NUM.findall(norm_text(s))]


def fmt_dim(v):
    if v is None:
        return ""
    return str(int(v)) if abs(v - round(v)) < 1e-9 else f"{v:g}".replace(".", ",")


def price_from_text(s):
    """'13500 грн (стара ціна 15600 грн...)' -> 13500.0 (беремо поточну)."""
    s = norm_text(s)
    if not s:
        return None
    cut = re.search(r"стара ціна", s, re.IGNORECASE)
    if cut:
        s = s[: cut.start()]
    m = re.search(r"([\d\s]+(?:[.,]\d+)?)", s)
    return to_num(m.group(1)) if m else None


def mm(v, unit):
    return v * 1000 if v is not None and unit.strip().lower() == "м" else v


def detect_type(name, section=""):
    t = norm_text(f"{name} {section}").lower()
    if "тканин" in t:
        return "Тканина"
    if "весло" in t or "лопать" in t:
        return "Весло"
    if "валіз" in t:
        return "Валіза"
    if "телескоп" in t:
        return "Телескопічна трубка"
    if "тюнінг" in t or "послуга" in t:
        return "Послуга"
    if "з'єдн" in t or "зєдн" in t or "з’єдн" in t:
        return "З'єднання з отворами" if "отвор" in t else "З'єднання труб"
    if "прямокутн" in t and "труб" in t:
        return "Труба прямокутна"
    if "авіамодел" in t or "кесон" in t:
        return "Пластина авіамодельна"
    if ("квадрат" in t and ("пол" in t or "труб" in t)) or "пола рейка" in t \
            or "труба квадратна" in t or "трубка квадратна" in t:
        return "Труба квадратна"
    if "квадратний стрижень" in t or "стрижень квадратний" in t:
        return "Стрижень квадратний"
    if "рейка" in t:
        return "Рейка"
    if "пруток" in t or "стрижень" in t:
        return "Пруток"
    if "труб" in t:
        return "Труба кругла"
    if "плит" in t or "пластин" in t or "лист" in t:
        return "Лист / пластина"
    return "Інше"


def detect_tech(*texts):
    """Технологія — тільки те, що прямо написане в джерелі."""
    t = norm_text(" ".join(str(x or "") for x in texts)).lower()
    if "пултруз" in t or "pultr" in t:
        return "Пултрузія"
    if "twill" in t or "твіл" in t:
        return "3K Twill"
    if "plaine" in t or "плейн" in t or "plain" in t:
        return "3K Plaine"
    return ""


def tech_group(tech):
    """Пултрузія і тканий карбон — різні товари. Twill vs Plaine — лише
    плетіння, для порівняння цін це одна група."""
    if tech == "Пултрузія":
        return "Пултрузія"
    if tech in ("3K Twill", "3K Plaine"):
        return "Тканий 3K"
    return ""


def canon(ptype, dims):
    """Канонічний ключ (тип + розміри) і людський підпис розміру."""
    d = [x for x in dims if x is not None]
    if ptype == "Труба кругла" and len(d) >= 2:
        outer, inner = max(d[0], d[1]), min(d[0], d[1])
        length = d[2] if len(d) > 2 else 1000.0
        return (ptype, round(outer, 2), round(inner, 2), round(length, 1)), \
               f"⌀{fmt_dim(outer)}×⌀{fmt_dim(inner)}×{fmt_dim(length)}"
    if ptype == "Труба квадратна" and len(d) >= 3:
        if len(d) >= 4:
            a, inner, length = max(d[0], d[1]), d[2], d[3]
        else:
            a, inner, length = max(d[0], d[1]), None, d[2]
        mid = f"×{fmt_dim(inner)}" if inner else ""
        return (ptype, round(a, 2), round(inner, 2) if inner else None, round(length, 1)), \
               f"{fmt_dim(a)}×{fmt_dim(a)}{mid}×{fmt_dim(length)}"
    if ptype == "Труба прямокутна" and len(d) >= 3:
        a, b = max(d[0], d[1]), min(d[0], d[1])
        return (ptype, round(a, 2), round(b, 2), round(d[2], 1)), \
               f"{fmt_dim(a)}×{fmt_dim(b)}×{fmt_dim(d[2])}"
    if ptype == "Стрижень квадратний" and len(d) >= 2:
        a = max(d[0], d[1])
        length = d[2] if len(d) > 2 else 1000.0
        return (ptype, round(a, 2), round(length, 1)), \
               f"{fmt_dim(a)}×{fmt_dim(a)}×{fmt_dim(length)}"
    if ptype == "Пруток" and len(d) >= 1:
        length = d[1] if len(d) > 1 else 1000.0
        return (ptype, round(d[0], 2), round(length, 1)), \
               f"⌀{fmt_dim(d[0])}×{fmt_dim(length)}"
    if ptype == "Рейка" and len(d) >= 2:
        w, th = max(d[0], d[1]), min(d[0], d[1])
        length = d[2] if len(d) > 2 else 1000.0
        return (ptype, round(w, 2), round(th, 2), round(length, 1)), \
               f"{fmt_dim(w)}×{fmt_dim(th)}×{fmt_dim(length)}"
    if ptype in ("Лист / пластина", "Пластина авіамодельна") and len(d) >= 3:
        o = sorted(d)
        th, rest = o[0], sorted(o[1:], reverse=True)
        return (ptype, round(th, 2), round(rest[0], 1), round(rest[1], 1)), \
               f"{fmt_dim(th)}×{fmt_dim(rest[0])}×{fmt_dim(rest[1])}"
    if ptype in ("З'єднання труб", "З'єднання з отворами") and len(d) >= 1:
        return (ptype, round(d[0], 2)), f"⌀{fmt_dim(d[0])}"
    return None, ""


# ------------------------------------------------------------ завантажувачі

def load_our(path, tag):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    ix = {norm_text(h): i for i, h in enumerate(rows[0])}
    out = []
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        section = norm_text(r[ix["Раздел"]])
        name = norm_text(r[ix["Название (UA)"]])
        if "арбон" not in f"{section} {name}".lower():
            continue
        ptype = detect_type(name, section)
        # Технологія: спершу назва — вона авторитетна. Опис беремо тільки якщо
        # в назві про технологію нічого немає (у Leti так у більшості позицій).
        # Описи бувають скопійовані з іншого товару: напр. у FlyShop
        # «Труба карбонова 3k Twill 6х4х1000» має опис про пултрузію.
        tech = (detect_tech(name)
                or detect_tech(r[ix["Описание товара (UA)"]], r[ix["Короткое описание (UA)"]]))
        src = re.sub(r"3\s*[kк]\s*(twill|plaine)?", " ", name, flags=re.IGNORECASE)
        src = DIAM.sub(" ", re.sub(r"\(.*?\)", " ", src))
        key, size = canon(ptype, nums_in(src))
        qty = r[ix["Количество"]]
        out.append(dict(
            src=tag, article=norm_text(r[0]), name=name, section=section,
            ptype=ptype, key=key, size=size, tech=tech, tech_group=tech_group(tech),
            price=to_num(r[ix["Цена"]]), currency=norm_text(r[ix["Валюта"]]) or "UAH",
            avail=norm_text(r[ix["Наличие"]]),
            qty=qty if isinstance(qty, (int, float)) else to_num(qty),
            url=norm_text(r[ix["Ссылка"]]),
        ))
    return out


def load_carbon_net(path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    out = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        name, art, price_txt, avail, specs, _d, cat, url = [norm_text(x) for x in r[:8]]
        if not name:
            continue
        ptype = detect_type(name, cat)
        tech = detect_tech(name, cat, specs)
        outer = re.search(r"Зовнішній\s*[∅Øø]?\s*([\d.,]+)", specs)
        inner = re.search(r"Внутрішній\s*[∅Øø]?\s*([\d.,]+)", specs)
        diam = re.search(r"Діаметр\s*[∅Øø]?\s*([\d.,]+)", specs)
        length = re.search(r"Довжина\s*([\d.,]+)\s*(мм|м)?", specs)
        size_sp = re.search(r"Розмір\s*([\d.,]+)\s*[×xх]\s*([\d.,]+)", specs)
        L = mm(to_num(length.group(1)), length.group(2) or "") if length else 1000.0
        if ptype == "Труба кругла" and outer and inner:
            dims = [to_num(outer.group(1)), to_num(inner.group(1)), L]
        elif ptype == "Пруток" and diam:
            dims = [to_num(diam.group(1)), L]
        elif ptype in ("З'єднання труб", "З'єднання з отворами") and inner:
            dims = [to_num(inner.group(1))]
        elif ptype == "Лист / пластина" and size_sp:
            th = nums_in(re.sub(r"\(.*?\)", " ", name))
            dims = [th[0] if th else None, to_num(size_sp.group(1)), to_num(size_sp.group(2))]
        else:
            s = re.sub(r"3\s*[kК]\s*twill", " ", re.sub(r"\(.*?\)", " ", name), flags=re.IGNORECASE)
            dims = nums_in(DIAM.sub(" ", s))
            wall = re.search(r"стінка\s*([\d.,]+)", name)
            if ptype == "Труба квадратна" and len(dims) == 3 and wall:
                dims = [dims[0], dims[1], dims[0] - 2 * to_num(wall.group(1)), dims[2]]
        key, size = canon(ptype, dims)
        out.append(dict(
            src="CN", article=art, name=name, section=cat, ptype=ptype, key=key, size=size,
            tech=tech, tech_group=tech_group(tech), price=price_from_text(price_txt),
            currency="UAH", avail=avail, qty=None, url=url,
        ))
    return out


def load_com_posit(path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    out = []
    for r in list(ws.iter_rows(values_only=True))[1:]:
        name, price_txt, cur, avail, specs, url = [norm_text(x) for x in r[:6]]
        if not name:
            continue
        group = url.split("carbon-products/")[-1].split("/")[0] if url else ""
        ptype = detect_type(name, group)
        tech = detect_tech(name, specs)
        io_ = re.search(r"внутр/зовн[:\s]*([\d.,]+)\s*/\s*([\d.,]+)", specs, re.IGNORECASE)
        length = re.search(r"Довжина[:\s]*([\d.,]+)\s*(мм|м)?", specs, re.IGNORECASE)
        th = re.search(r"Товщина[:\s]*([\d.,]+)", specs, re.IGNORECASE)
        size_sp = re.search(r"Розмір[:\s]*([\d.,]+)\s*[×xх*]\s*([\d.,]+)", specs, re.IGNORECASE)
        if ptype == "Труба кругла" and io_:
            L = mm(to_num(length.group(1)), length.group(2) or "") if length else 1000.0
            dims = [to_num(io_.group(2)), to_num(io_.group(1)), L]
        elif ptype in ("Лист / пластина", "Пластина авіамодельна") and th and size_sp:
            dims = [to_num(th.group(1)), to_num(size_sp.group(1)), to_num(size_sp.group(2))]
        else:
            dims = nums_in(re.sub(r"\(.*?\)", " ", name))
        key, size = canon(ptype, dims)
        out.append(dict(
            src="CP", article="", name=name, section=group, ptype=ptype, key=key, size=size,
            tech=tech, tech_group=tech_group(tech), price=to_num(price_txt),
            currency=cur or "UAH", avail=avail, qty=None, url=url, price_raw=price_txt,
        ))
    return out


# ------------------------------------------------------------------ зіставлення

def build_rows(items):
    """Групує позиції за (тип+розмір+технологія).

    Технологія береться буквально. Якщо джерело її не вказало, позиція
    приєднується до маркованої групи ЛИШЕ коли вибору немає: у цьому розмірі
    існує рівно одна технологічна група і в цьому джерелі рівно одна позиція
    такого розміру. Інакше лишається окремим рядком «не вказано».
    """
    buckets, no_dims = defaultdict(list), []
    for it in items:
        (no_dims if it["key"] is None else buckets[it["key"]]).append(it)

    rows, dupes = {}, []
    for key, group in buckets.items():
        groups = {it["tech_group"] for it in group if it["tech_group"]}
        for it in group:
            tg = it["tech_group"]
            if not tg:
                same_src = [x for x in group if x["src"] == it["src"]]
                tg = list(groups)[0] if len(groups) == 1 and len(same_src) == 1 else "не вказано"
            slot = rows.setdefault((key, tg), {})
            if it["src"] in slot:
                dupes.append(it)
            else:
                slot[it["src"]] = it
    return rows, no_dims, dupes


# ------------------------------------------------------------------- запис

HEAD = PatternFill("solid", fgColor="1F3864")
GREY = PatternFill("solid", fgColor="F2F2F2")


def style_header(ws, cols, freeze):
    for i, (t, w) in enumerate(cols, 1):
        c = ws.cell(1, i, t)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}1"
    ws.row_dimensions[1].height = 30


def main():
    items = (load_our(LETI_FILE, "LETI") + load_our(FLY_FILE, "FLY")
             + load_carbon_net(CN_FILE) + load_com_posit(CP_FILE))
    rows, no_dims, dupes = build_rows(items)

    # com-posit має групи в EUR (весла, авіамодельні пластини). Вони не
    # збігаються з жодною нашою позицією, тож у цінову таблицю не йдуть —
    # інакше довелось би вигадувати курс. Виносимо окремо.
    other_cur = []
    for k, slot in list(rows.items()):
        cp = slot.get("CP")
        if cp and cp["currency"] != "UAH":
            other_cur.append(cp)
            del slot["CP"]
            if not slot:
                del rows[k]

    wb = openpyxl.Workbook()

    # ---------------------------------------------------- лист «Порівняння»
    ws = wb.active
    ws.title = "Порівняння"
    cols = [
        ("Тип", 22), ("Розмір, мм", 20), ("Технологія", 13), ("Назва", 44),
        ("Leti, грн", 11), ("FlyShop, грн", 12),
        ("carbon.net.ua, грн", 15), ("com-posit, грн", 14),
        ("Наявність Leti", 15), ("Наявність FlyShop", 16),
        ("Наявність carbon.net.ua", 17), ("Наявність com-posit", 15),
        ("Залишок Leti", 11), ("Залишок FlyShop", 12),
        ("Мін. конкурент", 13), ("Наша мін.", 11), ("Δ грн", 10), ("Δ %", 9),
        ("Де є", 30),
    ]
    style_header(ws, cols, "C2")

    ORDER = {t: i for i, t in enumerate([
        "Труба кругла", "Труба квадратна", "Труба прямокутна", "Стрижень квадратний",
        "Пруток", "Рейка", "Лист / пластина", "Пластина авіамодельна",
        "З'єднання труб", "З'єднання з отворами"])}

    def sort_key(kv):
        (key, tg), _ = kv
        return (ORDER.get(key[0], 99),
                *[x if isinstance(x, (int, float)) else 0 for x in key[1:]], tg)

    money = "#,##0.00"
    r = 2
    for (key, tg), slot in sorted(rows.items(), key=sort_key):
        L, F, C, P = (slot.get(s) for s in SRC)
        any_it = next(x for x in (L, F, C, P) if x)
        ws.cell(r, 1, key[0])
        ws.cell(r, 2, any_it["size"])
        ws.cell(r, 3, tg)
        ws.cell(r, 4, (L or F or C or P)["name"])
        for col, it in ((5, L), (6, F), (7, C), (8, P)):
            if it:
                ws.cell(r, col, it["price"]).number_format = money
        for col, it in ((9, L), (10, F), (11, C), (12, P)):
            if it:
                ws.cell(r, col, it["avail"])
        if L:
            ws.cell(r, 13, L["qty"])
        if F:
            ws.cell(r, 14, F["qty"])
        # мінімуми: явні IF, бо MIN(IF(...)) на порожній комірці дає 0
        ws.cell(r, 15, f'=IF(AND(G{r}="",H{r}=""),"",'
                       f'IF(G{r}="",H{r},IF(H{r}="",G{r},MIN(G{r},H{r}))))').number_format = money
        ws.cell(r, 16, f'=IF(AND(E{r}="",F{r}=""),"",'
                       f'IF(E{r}="",F{r},IF(F{r}="",E{r},MIN(E{r},F{r}))))').number_format = money
        ws.cell(r, 17, f'=IF(OR(O{r}="",P{r}=""),"",P{r}-O{r})').number_format = money
        ws.cell(r, 18, f'=IF(OR(O{r}="",P{r}="",O{r}=0),"",(P{r}-O{r})/O{r})').number_format = "0%"
        ws.cell(r, 19, " + ".join(SRC_TITLE[s] for s in SRC if slot.get(s)))
        r += 1

    rng = f"A2:{get_column_letter(len(cols))}{r-1}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND($Q2<>"",$Q2>0)'], fill=PatternFill("solid", fgColor="FCE4E4")))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'AND($Q2<>"",$Q2<0)'], fill=PatternFill("solid", fgColor="E2EFDA")))

    # ------------------------------------------- лист «Перевірка» (аудит)
    ws2 = wb.create_sheet("Перевірка")
    cols2 = [("Тип", 22), ("Розмір, мм", 18), ("Технологія", 13),
             ("Назва Leti", 40), ("Назва FlyShop", 40),
             ("Назва carbon.net.ua", 40), ("Назва com-posit", 40),
             ("Артикул Leti", 16), ("Артикул FlyShop", 16),
             ("URL Leti", 34), ("URL FlyShop", 34),
             ("URL carbon.net.ua", 34), ("URL com-posit", 34)]
    style_header(ws2, cols2, "C2")
    r2 = 2
    for (key, tg), slot in sorted(rows.items(), key=sort_key):
        L, F, C, P = (slot.get(s) for s in SRC)
        ws2.cell(r2, 1, key[0])
        ws2.cell(r2, 2, next(x for x in (L, F, C, P) if x)["size"])
        ws2.cell(r2, 3, tg)
        for col, it in ((4, L), (5, F), (6, C), (7, P)):
            if it:
                ws2.cell(r2, col, it["name"])
        for col, it in ((8, L), (9, F)):
            if it:
                ws2.cell(r2, col, it["article"])
        for col, it in ((10, L), (11, F), (12, C), (13, P)):
            if it and it.get("url"):
                ws2.cell(r2, col, it["url"])
        r2 += 1

    # --------------------------------- лист «Не порівнювалось» (нічого не губимо)
    ws3 = wb.create_sheet("Не порівнювалось")
    cols3 = [("Джерело", 17), ("Причина", 38), ("Тип", 22), ("Назва", 60),
             ("Ціна", 12), ("Валюта", 9), ("Наявність", 18), ("URL", 38)]
    style_header(ws3, cols3, "A2")
    left = ([dict(x, _why="у назві немає розмірів") for x in no_dims]
            + [dict(x, _why="дубль такого ж розміру в тому самому магазині") for x in dupes]
            + [dict(x, _why="ціна не в гривні — курс не вигадуємо") for x in other_cur])
    for i, it in enumerate(sorted(left, key=lambda x: (x["src"], x["ptype"], x["name"])), 2):
        ws3.cell(i, 1, SRC_TITLE[it["src"]])
        ws3.cell(i, 2, it["_why"])
        ws3.cell(i, 3, it["ptype"])
        ws3.cell(i, 4, it["name"])
        ws3.cell(i, 5, it["price"] if it["price"] is not None else it.get("price_raw", ""))
        ws3.cell(i, 6, it["currency"])
        ws3.cell(i, 7, it["avail"])
        ws3.cell(i, 8, it.get("url", ""))

    # ------------------------------------------------------- лист «Як читати»
    ws4 = wb.create_sheet("Як читати", 0)
    ws4.column_dimensions["A"].width = 110
    txt = [
        ("Порівняння карбону: Leti, FlyShop, carbon.net.ua, com-posit.com.ua", True),
        ("", False),
        ("Лист «Порівняння» — головний: один рядок = одна позиція, поруч ціни всіх "
         "чотирьох сайтів.", False),
        ("Червоний рядок — ми дорожчі за конкурента, зелений — дешевші.", False),
        ("Порожня клітинка ціни означає, що такої позиції в цьому магазині немає.", False),
        ("", False),
        ("Звідки беруться цифри", True),
        ("Усі ціни, наявність і залишки скопійовані з файлів-джерел як є. "
         "Нічого не перераховано і не домислено.", False),
        ("Артикули в чотирьох джерелах різні, тому позиції зіставлені за типом виробу, "
         "розмірами в мм і технологією.", False),
        ("", False),
        ("Чому в таблиці є колонка «Технологія»", True),
        ("За одного розміру пултрузія і тканий карбон 3K — різні товари, різниця в ціні "
         "буває 2-3 рази (напр. труба ⌀10×8×1000: пултрузія 381 грн, 3K Twill 1045 грн).", False),
        ("Технологія прочитана з назви, опису та характеристик. Де в джерелі про неї "
         "не сказано і визначити однозначно не можна — стоїть «не вказано», і така "
         "позиція винесена окремим рядком, а не приліплена до чужої ціни.", False),
        ("", False),
        ("Лист «Перевірка» — ті самі рядки, але з назвами, артикулами й посиланнями "
         "кожного сайту, щоб можна було звірити, що зіставилось саме те.", False),
        ("Лист «Не порівнювалось» — позиції, які свідомо не потрапили в порівняння, "
         "із зазначенням причини.", False),
    ]
    for i, (line, bold) in enumerate(txt, 1):
        c = ws4.cell(i, 1, line)
        c.font = Font(bold=bold, size=12 if bold and i == 1 else 11)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(OUT_FILE)

    # ------------------------------------------------------------- звіт
    combos = defaultdict(int)
    for slot in rows.values():
        combos[" + ".join(SRC_TITLE[s] for s in SRC if slot.get(s))] += 1
    print(f"Збережено: {OUT_FILE}")
    print(f"  Порівняння: {r-2} рядків")
    for k, v in sorted(combos.items(), key=lambda x: -x[1]):
        print(f"     {v:5}  {k}")
    print(f"  Не порівнювалось: {len(left)} "
          f"(без розмірів {len(no_dims)}, дублі {len(dupes)}, не в грн {len(other_cur)})")


if __name__ == "__main__":
    main()
