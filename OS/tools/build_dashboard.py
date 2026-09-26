"""Собирает главный дашборд ОС «Мегаполис» (07_Дашборд.html) из файлов OS/*.md.

Единственный источник содержания — markdown-файлы ОС. После правки любого файла
запустить:  python3 OS/tools/build_dashboard.py
"""
import json, re, pathlib, html
import markdown

ROOT = pathlib.Path(__file__).resolve().parents[2]
OS = ROOT / "OS"
TPL = pathlib.Path(__file__).with_name("dashboard_template.html")
OUT = ROOT / "07_Дашборд.html"
TASKS = pathlib.Path(__file__).with_name("tasks_background.json")

TABS = [
    ("home", "Главная", None),
    ("model", "Направления", "B_Операционная_модель.md"),
    ("plan", "90 дней", "D_План_90_дней.md"),
    ("money", "Деньги", "E_Денежный_маршрут.md"),
    ("strategy", "Стратегия", "C_Стратегия_12_месяцев.md"),
    ("marketing", "Маркетинг", "F_Архитектура_маркетинга.md"),
    ("funnel", "Воронка", "G_Клиентский_путь_и_воронка.md"),
    ("roles", "Люди и AI", "H_Матрица_ролей.md"),
    ("decisions", "Решения", "J_Реестр_пробелов_и_решений.md"),
    ("monday", "Понедельник", "K_С_чего_начать_в_понедельник.md"),
    ("check", "Согласованность", "L_Проверка_согласованности.md"),
    ("map", "Карта (текст)", "A_Карта_бизнеса.md"),
    ("tasks", "Микрошаги (фон)", None),
]

def md_to_html(text):
    # mermaid-блоки → <pre class="mermaid">
    def repl(m):
        return '<pre class="mermaid">' + html.escape(m.group(1)) + '</pre>'
    text = re.sub(r"```mermaid\n(.*?)```", repl, text, flags=re.S)
    return markdown.markdown(text, extensions=["tables", "fenced_code", "sane_lists", "toc"])

def parse_table(text, header_start):
    """Возвращает строки первой таблицы, чей заголовок начинается с header_start."""
    lines = text.splitlines()
    for i, l in enumerate(lines):
        if l.startswith("| " + header_start):
            rows = []
            for l2 in lines[i + 2:]:
                if not l2.startswith("|"):
                    break
                rows.append([c.strip() for c in l2.strip().strip("|").split("|")])
            return rows
    return []

def clean(s):
    return re.sub(r"\*\*|`", "", s)

def main():
    pages = {}
    for key, _, fname in TABS:
        if fname:
            src = (OS / fname).read_text(encoding="utf-8")
            if key == "map":  # схема живёт на Главной, здесь только таблицы
                src = re.sub(r"```mermaid.*?```", "_Схема связей — на вкладке «Главная»._", src, flags=re.S)
            pages[key] = md_to_html(src)
    a = (OS / "A_Карта_бизнеса.md").read_text(encoding="utf-8")
    directions = [dict(n=r[0], name=clean(r[1]), result=clean(r[2]), owner=r[3], now=clean(r[4]),
                       state=clean(r[5]), first=clean(r[6])) for r in parse_table(a, "#")]
    j = (OS / "J_Реестр_пробелов_и_решений.md").read_text(encoding="utf-8")
    decisions = [dict(id=r[0], q=clean(r[1]), default=clean(r[4]), type=r[5], blocks=clean(r[6]))
                 for r in parse_table(j, "ID") if r[0].startswith("J-")]
    k = (OS / "K_С_чего_начать_в_понедельник.md").read_text(encoding="utf-8")
    monday = [dict(n=r[0], what=clean(r[1]), how=clean(r[3]), time=r[4], done=clean(r[5]))
              for r in parse_table(k, "#")]
    d = (OS / "D_План_90_дней.md").read_text(encoding="utf-8")
    stages = []
    for m in re.finditer(r"## (Этап [^\n]+)\n(.*?)(?=\n## |\Z)", d, flags=re.S):
        rows = [r for r in parse_table(m.group(2), "#")]
        gate = re.findall(r"\*\*(Условие перехода[^*]*|Контрольная точка[^*]*)\*\*([^\n]*)", m.group(2))
        stages.append(dict(title=clean(m.group(1)), ai=False,
                           steps=[dict(id=r[0], what=clean(r[1]), who=clean(r[2]), dep=clean(r[3])) for r in rows],
                           gates=[clean(g[0]) + clean(g[1]) for g in gate]))
    m = re.search(r"## (Параллельный трек[^\n]+)\n(.*?)(?=\n## |\Z)", d, flags=re.S)
    if m:
        rows = parse_table(m.group(2), "Когда")
        stages.append(dict(title=clean(m.group(1)), ai=True, gates=[],
                           steps=[dict(id="AI.%d" % (i + 1), what=clean(r[1]), who=clean(r[2]), dep=clean(r[0]))
                                  for i, r in enumerate(rows)]))
    money = {}
    try:
        from pycel import ExcelCompiler
        xl = ExcelCompiler(filename=str(OS / "I_Финансовая_модель.xlsx"))
        cols = "BCDEFGHIJKLM"
        money = dict(
            end=[round(xl.evaluate("'Движение денег'!%s14" % c)) for c in cols],
            inflow=[round(xl.evaluate("'Движение денег'!%s12" % c)) for c in cols],
            out=[round(xl.evaluate("'Движение денег'!%s7" % c) + xl.evaluate("'Движение денег'!%s8" % c)) for c in cols],
            deals_year=[round(xl.evaluate("Сценарии!%s8" % c), 1) for c in "BCD"],
            cpq_max=round(xl.evaluate("'Экономика сделки'!C16")),
            per_deal=round(xl.evaluate("'Экономика сделки'!C8")),
            gap_month=xl.evaluate("'Движение денег'!B16"))
    except Exception as e:  # модель не посчиталась — дашборд работает без графика
        print("pycel:", e)
    data = dict(money=money, tabs=[[k_, t] for k_, t, _ in TABS], pages=pages, directions=directions,
                decisions=decisions, monday=monday, stages=stages,
                tasks=json.loads(TASKS.read_text(encoding="utf-8")))
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    mer = re.search(r"```mermaid\n(.*?)```", a, flags=re.S)
    out = TPL.read_text(encoding="utf-8").replace("__DATA__", payload)
    out = out.replace("__MERMAID__", html.escape(mer.group(1)) if mer else "")
    OUT.write_text(out, encoding="utf-8")
    print("OK", OUT, len(out), "байт;", len(directions), "направлений,", len(decisions), "решений,",
          sum(len(s["steps"]) for s in stages), "шагов плана")

if __name__ == "__main__":
    main()
