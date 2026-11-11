# -*- coding: utf-8 -*-
"""Build the green-bond press-release workbook."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.comments import Comment

OUT = r"C:\Users\Vladimir\Desktop\Green_bonds\green_bonds_press_releases.xlsx"

DARK = "1F4E3D"
LIGHT = "F2F7F4"
AMBER = "FFF2CC"
GREY = "EDEDED"
GREEN_FILL = "E2EFDA"

thin = Side(style="thin", color="C9D6CF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---------------------------------------------------------------- main data
# (issuer, date, url, category, issue, bond_type, debut, verification, comment)
ROWS = [
    # ---------------- developers ----------------
    ("ПАО «Группа ЛСР»", dt.date(2026, 11, 2),
     "https://www.lsrgroup.ru/pressroom/novosti/2026/gruppa-lsr-vyiplatila-doxod-po-pervomu-kuponu-birzhevyix-zelenyix-obligaczij-serii-002r-02.html",
     "Девелопер", "002Р-02, 5 млрд ₽", "Зелёные", "Нет", "Открыто напрямую",
     "Пресс-релиза о самом размещении (02.10.2026) на сайте нет. В таблице — ближайший релиз эмитента по этому выпуску: выплата первого купона 02.11.2026."),

    ("ООО «Сэтл Групп»", dt.date(2026, 9, 14),
     "https://setlgroup.ru/news/2944",
     "Девелопер", "002Р-08, 5 млрд ₽", "Зелёные", "Да", "Открыто напрямую",
     "Дебютный выпуск; размещение 11.09.2026, срок 1 200 дней, купон — КС + 2,85%."),

    ("ПАО «Глоракс» (GloraX)", dt.date(2026, 6, 29),
     "https://glorax.com/investors/press-center/press-releases/pao-gloraks-nachinaet-razmeshenie-pervogo-vypuska-zelenyh-obligacij",
     "Девелопер", "002Р-01, 6,1 млрд ₽", "Зелёные", "Да", "Открыто напрямую",
     "Первый зелёный выпуск в истории компании; в релизе указан предварительный объём 3 млрд ₽, итоговый — 6,1 млрд ₽."),

    ("ГК «Пионер»", dt.date(2026, 3, 11),
     "https://pioneer.ru/news/pioner-pervym-v-2026-godu-vypustil-zelenye-obligatsii-obemom-4-6-mlrd-rubley",
     "Девелопер", "002Р-03, 4,6 млрд ₽", "Зелёные", "Нет", "Открыто напрямую",
     "Первый зелёный выпуск девелопера в 2026 г. (не дебют компании). Дата — начало торгов/размещения по данным Мосбиржи: на самой странице дата в текстовом слое не отображается."),

    ("ООО «ЛЕГЕНДА» (LEGENDA)", dt.date(2023, 5, 26),
     "https://legenda-dom.ru/news/legenda-pervoy-iz-zastroyshchikov-zhiloy-nedvizhimosti-vypustila-zelenye-obligacii",
     "Девелопер", "3 млрд ₽, 4 года", "Зелёные", "Да", "Открыто напрямую",
     "Первый зелёный выпуск среди застройщиков жилой недвижимости; ставка — КС + 5%, но не более 20%."),

    ("ИЦБ ДОМ.РФ (Ипотечный агент ДОМ.РФ)", dt.date(2022, 12, 1),
     "https://xn--90an6b.xn--d1aqf.xn--p1ai/media/news/dom-rf-razmestil-pervyy-v-rossii-vypusk-zelenykh-ipotechnykh-obligatsiy-/",
     "Финансовый институт развития / ипотечные ценные бумаги", "БО-001P-26, 5,5 млрд ₽", "Зелёные", "Да", "Открыто напрямую",
     "Первый в России выпуск зелёных ипотечных облигаций; эмитент — Ипотечный агент ДОМ.РФ, поручитель — ДОМ.РФ."),

    # ---------------- other issuers ----------------
    ("ПАО «Россети»", dt.date(2026, 8, 27),
     "https://www.rosseti.ru/sustainable-development/news/rosseti-razmestili-debyutnyy-vypusk-zelenykh-obligatsiy/",
     "Электросетевой холдинг", "21,5 млрд ₽, 7 лет", "Зелёные", "Да", "Подтверждено веб-архивом (снимок 11.11.2026)",
     "Дебютный выпуск; сайт закрыт антибот-защитой, поэтому страница подтверждена архивной копией. Размещение завершено 24.08.2026."),

    ("ПАО «АФК Система»", dt.date(2026, 7, 6),
     "https://www.sistema.ru/press/pressreleases/afk-sistema-poluchila-premiyu-infraros-za-luchshiy-proekt-elektrotransportnoy-infrastruktury-",
     "Инвестиционный холдинг", "002Р-06 и 002Р-07, 6,5 млрд ₽", "Зелёные", "Да", "Открыто напрямую",
     "Отдельного релиза о размещении на сайте найти не удалось. Приведён релиз о премии «ИнфраРос», где эмитент подтверждает дебютный выпуск зелёных облигаций в декабре 2025 г."),

    ("АО «ГТЛК»", dt.date(2026, 3, 24),
     "https://www.gtlk.ru/press_room/news/gtlk-uspeshno-razmestila-vypusk-chetyrekhletnikh-esg-obligatsiy-serii-002p-12-na-30-mlrd-rubley/",
     "Лизинговая компания", "002P-12, 30 млрд ₽", "Устойчивое развитие (ESG)", "Нет", "Открыто напрямую",
     "Выпуск ESG-облигаций (облигации устойчивого развития), а не «зелёных» в узком смысле; размещение 23.03.2026."),

    ("ППК «Российский экологический оператор»", dt.date(2025, 10, 17),
     "https://reo.ru/tpost/2z90f3p9v1-rossiiskii-ekologicheskii-operator-vperv",
     "Экооператор", "17,2 млрд ₽", "Зелёные", "Да", "Открыто напрямую",
     "Анонс дебютного выпуска; средства направлены зелёным займом АО «НЭО» на КПО «Новосёлки». Фактическое размещение — 20.11.2025."),

    ("АО «Невский экологический оператор»", dt.date(2025, 10, 17),
     "https://spb-neo.ru/o-kompanii/news/nevskiy-ekologicheskiy-operator-podtverzhdaet-sootvetstvie-printsipam-ustoychivogo-razvitiya-uchastn/",
     "Экооператор", "4,168 млрд ₽", "Зелёные", "Нет", "Открыто напрямую",
     "Релиз о верификации выпуска «Эксперт РА», а не о самом размещении; выпуск размещён в пользу Санкт-Петербурга."),

    ("ВЭБ.РФ", dt.date(2023, 11, 22),
     "https://veb.ru/press-tsentr/57335/",
     "Банк развития", "ПБО-002Р-39", "Зелёные", "Нет", "Подтверждено веб-архивом (индекс новостей)",
     "Релиз о закрытии книги заявок на размещение зелёных облигаций; выпуск не дебютный."),

    ("Правительство Москвы", dt.date(2021, 5, 28),
     "https://www.mos.ru/news/item/91414073/",
     "Субъект РФ", "74-й выпуск, 70 млрд ₽", "Зелёные", "Да", "Открыто напрямую",
     "Первый зелёный выпуск субъекта РФ; купон 7,38%, срок 7 лет, спрос — более 86 млрд ₽."),

    ("Госкорпорация «Росатом» (АО «Атомэнергопром»)", dt.date(2021, 6, 25),
     "https://rosatom.ru/journalist/news/predpriyatie-rosatoma-uspeshno-razmestilo-pervye-zelenye-obligatsii/",
     "Госкорпорация", "001Р-01, 10 млрд ₽", "Зелёные", "Да", "Подтверждено веб-архивом (снимок 05.08.2021)",
     "Первый выпуск зелёных облигаций; средства — на рефинансирование программы «Ветроэнергетика». Эмитент выпуска — АО «Атомэнергопром»."),

    ("ПАО Сбербанк", dt.date(2021, 11, 12),
     "https://www.sberbank.ru/ru/press_center/all/article?newsID=df54ac19-c25b-4383-a764-41eaf1432828&blockID=1303&regionID=77&lang=ru&type=NEWS",
     "Банк", "25 млрд ₽, 2 года", "Зелёные", "Да", "Подтверждено веб-архивом (индекс пресс-центра)",
     "Дебютное размещение зелёных облигаций; сайт банка блокирует доступ из-за рубежа, URL подтверждён архивным индексом пресс-центра."),

    ("АО «Коммерческая недвижимость ФПК «Гарант-Инвест»", dt.date(2019, 12, 19),
     "https://www.com-real.ru/news/ao-kommercheskaya-nedvizhimost-fpk-garant-invest-razmestilo-debyutnyy-vypusk-zelenyh-obligaciy-na-summu-500-mln-rubley",
     "Коммерческая недвижимость", "001Р-06, 500 млн ₽", "Зелёные", "Да", "Открыто напрямую",
     "Дебютный зелёный выпуск, 3 года, купон 11,5%; размещение 17.12.2019, дата публикации — по ленте новостей сайта."),

    ("ПАО КБ «Центр-инвест»", dt.date(2019, 11, 12),
     "https://centrinvest.com/about/news/33466",
     "Банк", "250 млн ₽", "Зелёные", "Да", "Открыто напрямую",
     "Первый зелёный банковский выпуск в России; размещение с 15.11.2019. Ссылка — англоязычная версия сайта банка (русская страница с этим ID недоступна)."),
]

# ---------------------------------------------------------------- workbook
wb = Workbook()
ws = wb.active
ws.title = "Пресс-релизы"

headers = ["Наименование эмитента", "Дата выхода релиза", "Ссылка на релиз на сайте эмитента"]

for c, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=c, value=h)
    cell.font = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=DARK)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER

ws.row_dimensions[1].height = 36

for i, (issuer, date, url, *_rest) in enumerate(ROWS, start=2):
    c1 = ws.cell(row=i, column=1, value=issuer)
    c1.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    c1.font = Font(name="Calibri", size=10.5)
    c2 = ws.cell(row=i, column=2, value=date)
    c2.number_format = "DD.MM.YYYY"
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.font = Font(name="Calibri", size=10.5)

    c3 = ws.cell(row=i, column=3, value=url)
    c3.hyperlink = url
    c3.font = Font(name="Calibri", size=8.5, color="0563C1", underline="single")
    c3.alignment = Alignment(vertical="center", wrap_text=True, indent=1)

    for c in (c1, c2, c3):
        c.border = BORDER
        if i % 2 == 0:
            c.fill = PatternFill("solid", fgColor=LIGHT)

    # highlight developers block (first 6 rows)
    if i <= 7:
        for c in (c1, c2, c3):
            c.fill = PatternFill("solid", fgColor=GREEN_FILL)

    ws.row_dimensions[i].height = 52

NOTES = {
    "ПАО «Группа ЛСР»": "Пресс-релиза о размещении (02.10.2026) на сайте нет. Это релиз о выплате первого купона по зелёному выпуску 002Р-02.",
    "ГК «Пионер»": "Дата приведена по началу торгов/размещению (Мосбиржа, 11.03.2026): на самой странице дата в текстовом слое не отображается.",
    "ПАО «АФК Система»": "Отдельного релиза о размещении на сайте нет. Это релиз о премии «ИнфраРос», подтверждающий дебютный выпуск зелёных облигаций в декабре 2025 г.",
    "АО «ГТЛК»": "Выпуск ESG-облигаций (облигации устойчивого развития), а не «зелёных» в узком смысле.",
    "АО «Невский экологический оператор»": "Релиз о верификации выпуска, а не о самом размещении.",
}
for i, (issuer, date, url, *_rest) in enumerate(ROWS, start=2):
    if issuer in NOTES:
        ws.cell(row=i, column=1).comment = Comment(NOTES[issuer], "DSH")
        ws.cell(row=i, column=1).comment.width = 320
        ws.cell(row=i, column=1).comment.height = 110

ws.column_dimensions["A"].width = 34
ws.column_dimensions["B"].width = 15
ws.column_dimensions["C"].width = 76
ws.column_dimensions["D"].width = 2

ws.freeze_panes = "A2"
ws.auto_filter.ref = None
ws.sheet_view.zoomScale = 90
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:1"

# ---------------------------------------------------------------- details sheet
ws4 = wb.create_sheet("Дополнительно")
det_headers = ["Эмитент", "Категория", "Выпуск", "Тип инструмента", "Дебютный", "Способ проверки", "Комментарий"]
for c, h in enumerate(det_headers, 1):
    cell = ws4.cell(row=1, column=c, value=h)
    cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=DARK)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
ws4.row_dimensions[1].height = 34

for i, (issuer, date, url, category, issue, btype, debut, check, comment) in enumerate(ROWS, start=2):
    vals = [issuer, category, issue, btype, debut, check, comment]
    for c, v in enumerate(vals, 1):
        cell = ws4.cell(row=i, column=c, value=v)
        cell.border = BORDER
        cell.font = Font(name="Calibri", size=9.5)
        cell.alignment = Alignment(vertical="top", wrap_text=True,
                                   horizontal="center" if c == 5 else "left")
    if debut == "Да":
        ws4.cell(row=i, column=5).fill = PatternFill("solid", fgColor=GREEN_FILL)
        ws4.cell(row=i, column=5).font = Font(name="Calibri", size=9.5, bold=True)
    if btype.startswith("Устойчивое"):
        ws4.cell(row=i, column=4).fill = PatternFill("solid", fgColor=AMBER)
    ws4.row_dimensions[i].height = 56

for col, w in zip("ABCDEFG", (38, 30, 26, 22, 10, 32, 70)):
    ws4.column_dimensions[col].width = w
ws4.column_dimensions["H"].width = 2
ws4.freeze_panes = "A2"
ws4.auto_filter.ref = None
ws4.sheet_view.zoomScale = 90

# ---------------------------------------------------------------- developers sheet
ws2 = wb.create_sheet("Девелоперы")

dev_headers = ["Девелопер", "Статус", "Зелёный выпуск", "Дата размещения", "Ссылка", "Комментарий"]
for c, h in enumerate(dev_headers, 1):
    cell = ws2.cell(row=1, column=c, value=h)
    cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=DARK)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
ws2.row_dimensions[1].height = 34

DEVS = [
    ("ООО «ЛЕГЕНДА»", "Релиз найден", "3 млрд ₽, 2023", "26.05.2023",
     "https://legenda-dom.ru/news/legenda-pervoy-iz-zastroyshchikov-zhiloy-nedvizhimosti-vypustila-zelenye-obligacii",
     "Первый зелёный выпуск среди застройщиков жилья. В основной таблице."),
    ("ГК «Пионер»", "Релиз найден", "002Р-03, 4,6 млрд ₽", "11.03.2026",
     "https://pioneer.ru/news/pioner-pervym-v-2026-godu-vypustil-zelenye-obligatsii-obemom-4-6-mlrd-rubley",
     "Первый зелёный выпуск девелопера в 2026 г. В основной таблице."),
    ("ПАО «Глоракс» (GloraX)", "Релиз найден", "002Р-01, 6,1 млрд ₽", "10.07.2026",
     "https://glorax.com/investors/press-center/press-releases/pao-gloraks-nachinaet-razmeshenie-pervogo-vypuska-zelenyh-obligacij",
     "Дебютный выпуск. В основной таблице."),
    ("ООО «Сэтл Групп»", "Релиз найден", "002Р-08, 5 млрд ₽", "11.09.2026",
     "https://setlgroup.ru/news/2944",
     "Дебютный выпуск. В основной таблице."),
    ("ПАО «Группа ЛСР»", "Релиза о размещении нет", "002Р-02, 5 млрд ₽", "02.10.2026",
     "https://www.lsrgroup.ru/pressroom/novosti/2026/gruppa-lsr-vyiplatila-doxod-po-pervomu-kuponu-birzhevyix-zelenyix-obligaczij-serii-002r-02.html",
     "На сайте только релиз о выплате первого купона (02.11.2026); релиза о размещении нет. В основной таблице — купонный релиз."),
    ("Группа «Эталон» (АО «Эталон-Финанс»)", "Релиз не найден", "002Р-06", "сбор заявок 09.11.2026",
     "https://www.moex.com/ru/listing/securities-cards.aspx?code=4B02-06-55338-H-002P",
     "Сайт etalon.ru недоступен (DNS), на etalongroup.com релиза нет. Ссылка — карточка выпуска на Мосбирже."),
    ("ГК «Брусника»", "Релиз не найден", "002Р-07 (7,4 млрд ₽); 002Р-11", "24.03.2026; 28.10.2026",
     "https://www.moex.com/n98782?nt=104",
     "На сайте только карточки выпусков и релизы о купонах; релиза о зелёном размещении нет."),
    ("ПАО «ПИК»", "Зелёных выпусков нет", "—", "—", "",
     "Обычные биржевые и замещающие выпуски; зелёных/ESG-облигаций не выявлено."),
    ("ПАО «ГК «Самолет»", "Зелёных выпусков нет", "—", "—", "",
     "Обычные биржевые выпуски; зелёных/ESG-облигаций не выявлено."),
    ("ГК «А101»", "Зелёных выпусков нет", "—", "—", "",
     "Обычные выпуски БО-001Р-01…03; зелёных облигаций не выявлено."),
    ("ГК «Инград»", "Зелёных выпусков нет", "—", "—", "",
     "Обычные выпуски 001P-02/002P-02; зелёных облигаций не выявлено."),
    ("ГК «ФСК»", "Зелёных выпусков нет", "—", "—", "",
     "Зелёных облигаций не выявлено (не путать с ПАО «ФСК ЕЭС»)."),
]

for i, row in enumerate(DEVS, start=2):
    name, status, issue, ddate, link, comment = row
    vals = [name, status, issue, ddate, "", comment]
    for c, v in enumerate(vals, 1):
        cell = ws2.cell(row=i, column=c, value=v)
        cell.border = BORDER
        cell.font = Font(name="Calibri", size=10)
        cell.alignment = Alignment(vertical="center", wrap_text=True,
                                   indent=2 if c == 1 else 1)
    ws2.cell(row=i, column=2).font = Font(name="Calibri", size=10, bold=True)
    if status == "Релиз найден":
        ws2.cell(row=i, column=2).fill = PatternFill("solid", fgColor=GREEN_FILL)
    elif status in ("Релиза о размещении нет", "Релиз не найден"):
        ws2.cell(row=i, column=2).fill = PatternFill("solid", fgColor=AMBER)
    else:
        ws2.cell(row=i, column=2).fill = PatternFill("solid", fgColor=GREY)
    if link:
        cl = ws2.cell(row=i, column=5,
                      value="Карточка на Мосбирже →" if "moex.com" in link else "Открыть релиз →")
        cl.hyperlink = link
        cl.font = Font(name="Calibri", size=9.5, color="0563C1", underline="single")
        cl.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    ws2.row_dimensions[i].height = 54

for col, w in zip("ABCDEF", (30, 22, 24, 20, 22, 60)):
    ws2.column_dimensions[col].width = w
ws2.column_dimensions["G"].width = 2
ws2.freeze_panes = "A2"
ws2.auto_filter.ref = None
ws2.sheet_view.zoomScale = 90

# ---------------------------------------------------------------- method sheet
ws3 = wb.create_sheet("Методика")
ws3.sheet_view.showGridLines = False
ws3.column_dimensions["A"].width = 3
ws3.column_dimensions["B"].width = 120

title = ws3.cell(row=2, column=2, value="Зелёные облигации российских эмитентов: пресс-релизы")
title.font = Font(name="Calibri", size=18, bold=True, color=DARK)
ws3.row_dimensions[2].height = 28

sub = ws3.cell(row=3, column=2, value="Подборка пресс-релизов на официальных сайтах эмитентов · данные на 11.11.2026")
sub.font = Font(name="Calibri", size=11, italic=True, color="595959")

blocks = [
    ("Что в таблице", [
        "17 пресс-релизов 17 эмитентов: 6 девелоперов (включая ИЦБ ДОМ.РФ) и 11 эмитентов из других отраслей.",
        "Приоритет отдан дебютным выпускам: из 17 релизов 11 сообщают о первом/дебютном зелёном выпуске.",
        "Первые строки выделены зелёным фоном — это блок девелоперов.",
    ]),
    ("Критерии отбора", [
        "Ссылка ведёт на домен самого эмитента (или его официальный IR-раздел), а не на Мосбиржу, Cbonds, Rusbonds, AK&M и другие агрегаторы.",
        "Страница открыта напрямую либо подтверждена копией в веб-архиве: проверены заголовок, дата публикации и принадлежность домена.",
        "Дата в столбце 2 — дата публикации релиза; если на странице видна только дата размещения, это отмечено в листе «Дополнительно».",
    ]),
    ("Важные оговорки", [
        "ГТЛК — выпуск ESG-облигаций (облигации устойчивого развития), а не «зелёных» в узком смысле.",
        "НЭО — релиз о верификации выпуска, а не о самом размещении; РЭО — анонс дебютного выпуска.",
        "АФК «Система» — отдельного релиза о размещении на сайте нет; приведён релиз о премии, подтверждающий дебютный выпуск.",
        "ЛСР — релиза о размещении нет; приведён релиз о выплате первого купона по зелёному выпуску 002Р-02.",
        "Пионер — дата приведена по началу торгов/размещению (11.03.2026), на самой странице дата в текстовом слое не отображается.",
        "РЖД (2020) и «Лизинг-Трейд» (2019) в таблицу не включены: ссылку РЖД не удалось подтвердить, у «Лизинг-Трейда» зелёного выпуска не было.",
    ]),
    ("Что осталось за рамками", [
        "Девелоперы с зелёными выпусками, но без релиза на сайте: ЛСР (002Р-02), «Эталон» (002Р-06), «Брусника» (002Р-07, 002Р-11).",
        "Девелоперы без зелёных выпусков: ПИК, «Самолет», А101, «Инград», ФСК.",
        "Эти случаи разобраны на листе «Девелоперы».",
    ]),
]

r = 5
for head, lines in blocks:
    c = ws3.cell(row=r, column=2, value=head)
    c.font = Font(name="Calibri", size=12, bold=True, color=DARK)
    c.fill = PatternFill("solid", fgColor=GREEN_FILL)
    c.alignment = Alignment(vertical="center")
    ws3.row_dimensions[r].height = 24
    r += 1
    for line in lines:
        c = ws3.cell(row=r, column=2, value="•  " + line)
        c.font = Font(name="Calibri", size=10.5)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws3.row_dimensions[r].height = 30
        r += 1
    r += 1

src = ws3.cell(row=r, column=2,
               value="Источники: официальные сайты эмитентов, Московская биржа, «Эксперт РА», АКРА, НКР, ИНФРАГРИН, Cbonds. Ссылки проверены 11.11.2026.")
src.font = Font(name="Calibri", size=9, italic=True, color="7F7F7F")
src.alignment = Alignment(wrap_text=True, vertical="top")

wb.properties.title = "Пресс-релизы российских эмитентов о размещении зелёных облигаций"
wb.properties.subject = "Зелёные облигации, девелоперы, ESG"
wb.properties.creator = "DeepSeek Harness"
wb.properties.created = dt.datetime(2026, 11, 11, 21, 0)

wb.save(OUT)
print("saved", OUT, "rows:", len(ROWS))





