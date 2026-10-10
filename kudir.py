import sqlite3
from pathlib import Path
from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "database.db"

app = FastAPI(title="КУДиР — ИП УСН 6%")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


# ---------- База данных ----------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS income (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            op_date TEXT NOT NULL,
            amount REAL NOT NULL,
            counterparty TEXT,
            comment TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pay_date TEXT NOT NULL,
            category TEXT NOT NULL,
            amount REAL NOT NULL,
            for_year INTEGER NOT NULL,
            comment TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_date TEXT NOT NULL,
            title TEXT NOT NULL,
            amount REAL,
            category TEXT,
            done INTEGER DEFAULT 0,
            comment TEXT,
            sort_order INTEGER NOT NULL DEFAULT 0
        )
    """)

    conn.commit()
    conn.close()


def migrate_db():
    conn = get_db()
    cur = conn.cursor()

    income_cols = [r["name"] for r in cur.execute("PRAGMA table_info(income)").fetchall()]
    if "counterparty" not in income_cols:
        cur.execute("ALTER TABLE income ADD COLUMN counterparty TEXT")

    event_cols = [r["name"] for r in cur.execute("PRAGMA table_info(events)").fetchall()]
    if "sort_order" not in event_cols:
        cur.execute(
            "ALTER TABLE events ADD COLUMN sort_order INTEGER NOT NULL DEFAULT 0"
        )
        # Сохраняем текущий порядок внутри каждого дня (по id)
        dates = [
            r["event_date"]
            for r in cur.execute(
                "SELECT DISTINCT event_date FROM events ORDER BY event_date"
            ).fetchall()
        ]
        for d in dates:
            rows = cur.execute(
                "SELECT id FROM events WHERE event_date = ? ORDER BY id ASC", (d,)
            ).fetchall()
            for i, row in enumerate(rows):
                cur.execute(
                    "UPDATE events SET sort_order = ? WHERE id = ?", (i, row["id"])
                )

    conn.commit()
    conn.close()


def next_event_sort_order(conn, event_date: str) -> int:
    row = conn.execute(
        "SELECT COALESCE(MAX(sort_order), -1) AS m FROM events WHERE event_date = ?",
        (event_date,),
    ).fetchone()
    return int(row["m"]) + 1


init_db()
migrate_db()


# ---------- Предзаполнение календаря ----------
FIXED_2026 = 57390.0


def seed_events():
    """Заполняем календарь стартовыми событиями, если он пуст."""
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) AS c FROM events")
    if cur.fetchone()["c"] > 0:
        conn.close()
        return

    fixed_3m = round(FIXED_2026 / 12 * 3, 2)

    events = [
        # 25.12.2026
        ("2026-12-25",
         "Фикс. взнос за 2026 (окт–дек)",
         fixed_3m, "fixed",
         "57390 / 12 × 3 мес. Успеть до 28.12.2026.", 0),

        # 20.04.2027 — дела в нужном порядке
        ("2027-04-20",
         "Подать декларацию за 2026",
         None, "declaration",
         "Уменьшение налога на фикс. взносы. Успеть до 25.04.2027. Сначала декларация, потом налог!", 0),

        ("2027-04-20",
         "Уплатить 1% свыше 300К за 2026",
         None, "one_percent",
         "Платим заранее (дедлайн 01.07.2027). Уменьшает аванс 1 кв. 2027.", 1),

        ("2027-04-20",
         "Уплатить уменьшенный налог за 2026",
         None, "tax",
         "Налог 6% минус фикс. взносы за 2026. Успеть до 28.04.2027.", 2),

        ("2027-04-20",
         "Аванс УСН за 1 квартал 2027",
         None, "tax",
         "Уменьшен на уплаченный 1% свыше 300К за 2026.", 3),
    ]
    for e in events:
        cur.execute(
            "INSERT INTO events (event_date, title, amount, category, done, comment, sort_order) "
            "VALUES (?,?,?,?,0,?,?)", e
        )
    conn.commit()
    conn.close()


seed_events()


# ---------- Реквизиты для декларации ----------
TAXPAYER = {
    "inn": "090108917865",
    "oktmo": "91701000001",
    "tax_authority": "0900",  # код НО (первые 4 цифры ИНН)
    "period_code": "34",      # 34 — год; 21 — I кв.; 31 — полугодие; 33 — 9 мес.
    "period_name": "год",
}

# ---------- Расчёты ----------
LIMIT_300K = 300000.0


def calc_summary(year: int = 2026):
    conn = get_db()
    cur = conn.cursor()

    cur.execute(
        "SELECT COALESCE(SUM(amount),0) AS s FROM income WHERE op_date LIKE ?",
        (f"{year}%",)
    )
    income_year = cur.fetchone()["s"] or 0.0

    cur.execute(
        "SELECT category, COALESCE(SUM(amount),0) AS s FROM payments "
        "WHERE for_year = ? GROUP BY category", (year,)
    )
    paid = {row["category"]: row["s"] for row in cur.fetchall()}

    conn.close()

    fixed_paid = paid.get("fixed", 0.0)
    one_percent_paid = paid.get("one_percent", 0.0)
    advance_q1 = paid.get("tax_q1", 0.0)
    advance_h1 = paid.get("tax_h1", 0.0)
    advance_9m = paid.get("tax_9m", 0.0)
    tax_year_paid = paid.get("tax", 0.0)
    tax_paid = tax_year_paid + advance_q1 + advance_h1 + advance_9m

    excess = max(0.0, income_year - LIMIT_300K)
    one_percent_due = excess * 0.01
    tax_gross = income_year * 0.06

    reduction = fixed_paid + one_percent_paid
    tax_final = max(0.0, tax_gross - reduction)
    tax_remaining = max(0.0, tax_final - tax_paid)

    return {
        "year": year,
        "income_year": income_year,
        "excess": excess,
        "one_percent_due": one_percent_due,
        "one_percent_paid": one_percent_paid,
        "tax_gross": tax_gross,
        "fixed_paid": fixed_paid,
        "reduction": reduction,
        "tax_final": tax_final,
        "tax_paid": tax_paid,
        "tax_remaining": tax_remaining,
        "advance_q1": advance_q1,
        "advance_h1": advance_h1,
        "advance_9m": advance_9m,
    }


# ---------- Роуты ----------
@app.get("/", response_class=HTMLResponse)
def index(request: Request, year: int = 2026):
    summary = calc_summary(year)
    conn = get_db()
    incomes = conn.execute(
        "SELECT * FROM income ORDER BY op_date DESC LIMIT 50"
    ).fetchall()
    payments = conn.execute(
        "SELECT * FROM payments ORDER BY pay_date DESC, id DESC LIMIT 50"
    ).fetchall()
    events = conn.execute(
        "SELECT * FROM events ORDER BY event_date ASC, sort_order ASC, id ASC"
    ).fetchall()
    years = [r["y"] for r in conn.execute(
        "SELECT DISTINCT CAST(substr(op_date,1,4) AS INTEGER) AS y "
        "FROM income WHERE op_date IS NOT NULL ORDER BY y DESC"
    ).fetchall()]
    conn.close()

    if year not in years:
        years.insert(0, year)

    date_counts: dict[str, int] = {}
    for e in events:
        date_counts[e["event_date"]] = date_counts.get(e["event_date"], 0) + 1

    return templates.TemplateResponse("index.html", {
        "request": request,
        "summary": summary,
        "incomes": incomes,
        "payments": payments,
        "events": events,
        "date_counts": date_counts,
        "years": years,
        "current_year": year,
        "taxpayer": TAXPAYER,
    })


# ----- Доходы -----
@app.post("/income/add")
def add_income(
    op_date: str = Form(...),
    amount: float = Form(...),
    counterparty: str = Form(""),
    comment: str = Form(""),
):
    conn = get_db()
    conn.execute(
        "INSERT INTO income (op_date, amount, counterparty, comment) VALUES (?,?,?,?)",
        (op_date, amount, counterparty.strip(), comment.strip()),
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


@app.post("/income/edit/{iid}")
def edit_income(
    iid: int,
    op_date: str = Form(...),
    amount: float = Form(...),
    counterparty: str = Form(""),
    comment: str = Form(""),
):
    conn = get_db()
    conn.execute(
        "UPDATE income SET op_date=?, amount=?, counterparty=?, comment=? WHERE id=?",
        (op_date, amount, counterparty.strip(), comment.strip(), iid),
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


@app.post("/income/delete/{iid}")
def delete_income(iid: int):
    conn = get_db()
    conn.execute("DELETE FROM income WHERE id=?", (iid,))
    conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


# ----- Платежи (факт) -----
@app.post("/payment/add")
def add_payment(
    pay_date: str = Form(...),
    category: str = Form(...),
    amount: float = Form(...),
    for_year: int = Form(...),
    comment: str = Form(""),
):
    conn = get_db()
    conn.execute(
        "INSERT INTO payments (pay_date, category, amount, for_year, comment) "
        "VALUES (?,?,?,?,?)",
        (pay_date, category, amount, for_year, comment.strip()),
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


@app.post("/payment/edit/{pid}")
def edit_payment(
    pid: int,
    pay_date: str = Form(...),
    category: str = Form(...),
    amount: float = Form(...),
    for_year: int = Form(...),
    comment: str = Form(""),
):
    conn = get_db()
    conn.execute(
        "UPDATE payments SET pay_date=?, category=?, amount=?, for_year=?, comment=? "
        "WHERE id=?",
        (pay_date, category, amount, for_year, comment.strip(), pid),
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


@app.post("/payment/delete/{pid}")
def delete_payment(pid: int):
    conn = get_db()
    conn.execute("DELETE FROM payments WHERE id=?", (pid,))
    conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


# ----- События (план) -----
@app.post("/event/add")
def add_event(
    event_date: str = Form(...),
    title: str = Form(...),
    amount: str = Form(""),
    category: str = Form("other"),
    comment: str = Form(""),
):
    amt = float(amount) if amount.strip() else None
    conn = get_db()
    sort_order = next_event_sort_order(conn, event_date)
    conn.execute(
        "INSERT INTO events (event_date, title, amount, category, done, comment, sort_order) "
        "VALUES (?,?,?,?,0,?,?)",
        (event_date, title, amt, category, comment, sort_order),
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


@app.post("/event/edit/{eid}")
def edit_event(
    eid: int,
    event_date: str = Form(...),
    title: str = Form(...),
    amount: str = Form(""),
    category: str = Form("other"),
    comment: str = Form(""),
):
    amt = float(amount) if amount.strip() else None
    conn = get_db()
    row = conn.execute(
        "SELECT event_date, sort_order FROM events WHERE id=?", (eid,)
    ).fetchone()
    if not row:
        conn.close()
        return RedirectResponse("/", status_code=303)

    sort_order = row["sort_order"]
    if row["event_date"] != event_date:
        sort_order = next_event_sort_order(conn, event_date)

    conn.execute(
        "UPDATE events SET event_date=?, title=?, amount=?, category=?, comment=?, "
        "sort_order=? WHERE id=?",
        (event_date, title, amt, category, comment, sort_order, eid),
    )
    conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


@app.post("/event/toggle/{eid}")
def toggle_event(eid: int):
    conn = get_db()
    row = conn.execute("SELECT done FROM events WHERE id=?", (eid,)).fetchone()
    if row:
        conn.execute("UPDATE events SET done=? WHERE id=?",
                     (0 if row["done"] else 1, eid))
        conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


@app.post("/event/delete/{eid}")
def delete_event(eid: int):
    conn = get_db()
    conn.execute("DELETE FROM events WHERE id=?", (eid,))
    conn.commit()
    conn.close()
    return RedirectResponse("/", status_code=303)


@app.post("/event/reorder")
async def reorder_events(request: Request):
    """Сохранить порядок дел внутри одного дня. Тело: {\"ids\": [id1, id2, ...]}."""
    try:
        data = await request.json()
    except Exception:
        return JSONResponse({"ok": False, "error": "invalid json"}, status_code=400)

    ids = data.get("ids") if isinstance(data, dict) else None
    if not isinstance(ids, list) or not ids or not all(isinstance(x, int) for x in ids):
        return JSONResponse({"ok": False, "error": "ids required"}, status_code=400)

    conn = get_db()
    rows = conn.execute(
        f"SELECT id, event_date FROM events WHERE id IN ({','.join('?' * len(ids))})",
        ids,
    ).fetchall()
    by_id = {r["id"]: r["event_date"] for r in rows}
    if len(by_id) != len(ids) or len(set(by_id.values())) != 1:
        conn.close()
        return JSONResponse(
            {"ok": False, "error": "ids must belong to the same day"},
            status_code=400,
        )

    for i, eid in enumerate(ids):
        conn.execute("UPDATE events SET sort_order = ? WHERE id = ?", (i, eid))
    conn.commit()
    conn.close()
    return JSONResponse({"ok": True})

# ---------- Экспорт ----------
CATEGORY_RU = {
    "fixed": "Фикс. взносы",
    "one_percent": "1% свыше 300К",
    "tax": "Налог УСН (год)",
    "tax_q1": "Аванс УСН I кв.",
    "tax_h1": "Аванс УСН полугодие",
    "tax_9m": "Аванс УСН 9 мес.",
    "declaration": "Декларация",
    "other": "Другое",
}


def _fetch_all():
    conn = get_db()
    incomes = conn.execute(
        "SELECT * FROM income ORDER BY op_date ASC, id ASC"
    ).fetchall()
    payments = conn.execute(
        "SELECT * FROM payments ORDER BY pay_date ASC, id ASC"
    ).fetchall()
    events = conn.execute(
        "SELECT * FROM events ORDER BY event_date ASC, sort_order ASC, id ASC"
    ).fetchall()
    conn.close()
    return incomes, payments, events


def _summary_for_export(year: int):
    return calc_summary(year)


def _today():
    return datetime.now().strftime("%Y-%m-%d")


# ----- TXT -----
@app.get("/export/txt")
def export_txt(year: int = 2026):
    incomes, payments, events = _fetch_all()
    s = _summary_for_export(year)

    lines = []
    lines.append("=" * 70)
    lines.append(f"КУДиР — ИП УСН 6% Доходы. Экспорт от {_today()}")
    lines.append(f"Отчётный год: {year}")
    lines.append("=" * 70)
    lines.append("")

    def _money_or_none(amount):
        if amount:
            return f"{amount:>14,.2f} ₽"
        return f"{'не уплачен':>16}"

    # Реквизиты
    tp = TAXPAYER
    lines.append("--- РЕКВИЗИТЫ ДЛЯ ДЕКЛАРАЦИИ ---")
    lines.append(f"ИНН:                   {tp['inn']}")
    lines.append(f"ОКТМО:                 {tp['oktmo']}")
    lines.append(f"Код налогового органа: {tp['tax_authority']}")
    lines.append(f"Отчётный период:       {tp['period_code']} ({tp['period_name']}, {year})")
    lines.append("")

    # Сводка
    lines.append("--- СВОДКА ---")
    lines.append(f"Доходы за {year}:              {s['income_year']:>14,.2f} ₽")
    lines.append(f"Превышение 300 000:            {s['excess']:>14,.2f} ₽")
    lines.append(f"1% свыше 300К (к уплате):      {s['one_percent_due']:>14,.2f} ₽")
    lines.append(f"Налог 6% (брутто):             {s['tax_gross']:>14,.2f} ₽")
    lines.append(f"Уплачено фикс. взносов:        {s['fixed_paid']:>14,.2f} ₽")
    lines.append(f"Уплачено 1%:                   {s['one_percent_paid']:>14,.2f} ₽")
    lines.append(f"Уменьшение налога:             {s['reduction']:>14,.2f} ₽")
    lines.append(f"Налог к уплате:                {s['tax_final']:>14,.2f} ₽")
    lines.append(f"Уплачено налога (всего):       {s['tax_paid']:>14,.2f} ₽")
    lines.append(f"Налог остаток:                 {s['tax_remaining']:>14,.2f} ₽")
    lines.append("")

    # Авансы
    lines.append("--- АВАНСОВЫЕ ПЛАТЕЖИ УСН ---")
    lines.append(f"Аванс за I квартал:            {_money_or_none(s['advance_q1'])}")
    lines.append(f"Аванс за полугодие:            {_money_or_none(s['advance_h1'])}")
    lines.append(f"Аванс за 9 месяцев:            {_money_or_none(s['advance_9m'])}")
    lines.append("")

    # Доходы
    lines.append("--- ДОХОДЫ ---")
    lines.append(f"{'Дата':<12} {'Сумма':>14}  {'Контрагент':<25} Комментарий")
    lines.append("-" * 70)
    total_inc = 0.0
    for i in incomes:
        total_inc += i["amount"]
        lines.append(
            f"{i['op_date']:<12} {i['amount']:>14,.2f}  "
            f"{(i['counterparty'] or '—'):<25} {i['comment'] or ''}"
        )
    lines.append("-" * 70)
    lines.append(f"{'ИТОГО':<12} {total_inc:>14,.2f}")
    lines.append("")

    # Платежи
    lines.append("--- ПЛАТЕЖИ (факт) ---")
    lines.append(f"{'Дата':<12} {'Категория':<24} {'Сумма':>14}  {'За год':<8} Комментарий")
    lines.append("-" * 70)
    total_pay = 0.0
    for p in payments:
        total_pay += p["amount"]
        cat = CATEGORY_RU.get(p["category"], p["category"])
        lines.append(
            f"{p['pay_date']:<12} {cat:<24} {p['amount']:>14,.2f}  "
            f"{p['for_year']:<8} {p['comment'] or ''}"
        )
    lines.append("-" * 70)
    lines.append(f"{'ИТОГО':<37} {total_pay:>14,.2f}")
    lines.append("")

    # Календарь
    lines.append("--- КАЛЕНДАРЬ (план) ---")
    lines.append(f"{'Дата':<12} {'Категория':<24} {'Сумма':>14}  {'Готово':<8} Дело")
    lines.append("-" * 70)
    for e in events:
        cat = CATEGORY_RU.get(e["category"], e["category"] or "—")
        amt = f"{e['amount']:,.2f}" if e["amount"] else ""
        done = "✔" if e["done"] else "○"
        lines.append(
            f"{e['event_date']:<12} {cat:<24} {amt:>14}  {done:<8} {e['title'] or ''}"
        )
        if e["comment"]:
            lines.append(f"{'':<12} {'':<24} {'':>14}  {'':<8}   └ {e['comment']}")
    lines.append("")
    lines.append("=" * 70)

    text = "\n".join(lines)
    buf = io.BytesIO(text.encode("utf-8-sig"))  # BOM для Windows Notepad
    fname = f"kudir_export_{year}_{_today()}.txt"
    return StreamingResponse(
        buf,
        media_type="text/plain; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )


# ----- XLSX -----
@app.get("/export/xlsx")
def export_xlsx(year: int = 2026):
    incomes, payments, events = _fetch_all()
    s = _summary_for_export(year)

    wb = Workbook()
    bold = Font(bold=True)
    head_fill = PatternFill("solid", fgColor="1E3A8A")
    head_font = Font(bold=True, color="FFFFFF")
    money_fmt = '#,##0.00 "₽"'

    def style_header(ws, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    def style_row_header(ws, row, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    def append_money_or_none(ws, name, amount):
        if amount:
            ws.append([name, amount])
            ws.cell(row=ws.max_row, column=2).number_format = money_fmt
        else:
            ws.append([name, "не уплачен"])

    # --- Лист «Сводка» ---
    ws = wb.active
    ws.title = "Сводка"
    tp = TAXPAYER
    ws.append(["Реквизит", "Значение"])
    style_header(ws, 2)
    ws.append(["ИНН", tp["inn"]])
    ws.append(["ОКТМО", tp["oktmo"]])
    ws.append(["Код налогового органа", tp["tax_authority"]])
    ws.append(["Отчётный период", f"{tp['period_code']} ({tp['period_name']}, {year})"])
    ws.append([])

    header_row = ws.max_row + 1
    ws.append(["Показатель", f"Значение ({year})"])
    style_row_header(ws, header_row, 2)

    rows = [
        ("Доходы", s["income_year"]),
        ("Превышение 300 000", s["excess"]),
        ("1% свыше 300К (к уплате)", s["one_percent_due"]),
        ("Налог 6% (брутто)", s["tax_gross"]),
        ("Уплачено фикс. взносов", s["fixed_paid"]),
        ("Уплачено 1%", s["one_percent_paid"]),
        ("Уменьшение налога", s["reduction"]),
        ("Налог к уплате", s["tax_final"]),
        ("Уплачено налога (всего)", s["tax_paid"]),
        ("Налог остаток", s["tax_remaining"]),
    ]
    for name, val in rows:
        ws.append([name, val])
        ws.cell(row=ws.max_row, column=2).number_format = money_fmt

    ws.append([])
    adv_header = ws.max_row + 1
    ws.append(["Авансовые платежи УСН", f"За {year}"])
    style_row_header(ws, adv_header, 2)
    append_money_or_none(ws, "Аванс за I квартал", s["advance_q1"])
    append_money_or_none(ws, "Аванс за полугодие", s["advance_h1"])
    append_money_or_none(ws, "Аванс за 9 месяцев", s["advance_9m"])

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 28

    # --- Лист «Доходы» ---
    ws = wb.create_sheet("Доходы")
    ws.append(["Дата", "Сумма", "Контрагент", "Комментарий"])
    style_header(ws, 4)
    for i in incomes:
        ws.append([i["op_date"], i["amount"], i["counterparty"] or "", i["comment"] or ""])
        ws.cell(row=ws.max_row, column=2).number_format = money_fmt
    # Итого
    ws.append(["ИТОГО", f"=SUM(B2:B{ws.max_row})", "", ""])
    ws.cell(row=ws.max_row, column=1).font = bold
    ws.cell(row=ws.max_row, column=2).font = bold
    ws.cell(row=ws.max_row, column=2).number_format = money_fmt
    for col, w in zip("ABCD", (14, 16, 28, 40)):
        ws.column_dimensions[col].width = w

    # --- Лист «Платежи» ---
    ws = wb.create_sheet("Платежи")
    ws.append(["Дата", "Категория", "Сумма", "За год", "Комментарий"])
    style_header(ws, 5)
    for p in payments:
        ws.append([
            p["pay_date"],
            CATEGORY_RU.get(p["category"], p["category"]),
            p["amount"],
            p["for_year"],
            p["comment"] or "",
        ])
        ws.cell(row=ws.max_row, column=3).number_format = money_fmt
    ws.append(["ИТОГО", "", f"=SUM(C2:C{ws.max_row})", "", ""])
    ws.cell(row=ws.max_row, column=1).font = bold
    ws.cell(row=ws.max_row, column=3).font = bold
    ws.cell(row=ws.max_row, column=3).number_format = money_fmt
    for col, w in zip("ABCDE", (14, 26, 16, 10, 40)):
        ws.column_dimensions[col].width = w

    # --- Лист «Календарь» ---
    ws = wb.create_sheet("Календарь")
    ws.append(["Дата", "Категория", "Дело", "Сумма", "Готово", "Комментарий"])
    style_header(ws, 6)
    for e in events:
        ws.append([
            e["event_date"],
            CATEGORY_RU.get(e["category"], e["category"] or ""),
            e["title"] or "",
            e["amount"] if e["amount"] else None,
            "Да" if e["done"] else "Нет",
            e["comment"] or "",
        ])
        if e["amount"]:
            ws.cell(row=ws.max_row, column=4).number_format = money_fmt
    for col, w in zip("ABCDEF", (14, 18, 40, 16, 10, 40)):
        ws.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"kudir_export_{year}_{_today()}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{fname}"'},
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("kudir:app", host="127.0.0.1", port=8000, reload=True)