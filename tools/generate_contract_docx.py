#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор docs/contract_v2.docx — Word-версия договора проката.

Дизайн-цели:
  * максимум 2 страницы A4 (узкие поля, 10pt текст, плотный интервал);
  * шапка Арендатора в таблице, чтобы поля не разъезжались;
  * линии под заполнение — через подчёркнутые пробелы одного рана
    (не разрываются переносами, ровные).

Без внешних зависимостей: docx собирается стандартным
zipfile + ручной Office Open XML.
"""

from __future__ import annotations

import os
import zipfile
from html import escape
from pathlib import Path

# --------------------------- параметры стилей ---------------------------

FONT_TEXT = "Times New Roman"
FONT_HEAD = "Arial"

# half-points (Word size = 2 * pt)
SIZE_TITLE = 28   # 14pt
SIZE_H2 = 24      # 12pt
SIZE_H3 = 22      # 11pt
SIZE_TEXT = 22    # 11pt
SIZE_NOTE = 20    # 10pt

# spacing — twips (1/20 pt)
SPACE_AFTER = 60
SPACE_AFTER_HEAD = 100
LINE_SPACING = 240   # 1.0

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

# --------------------------- runs / paragraphs ---------------------------


def run(text, *, bold=False, italic=False, font=FONT_TEXT,
        size=SIZE_TEXT, color=None, underline=False):
    rpr = ['<w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}"/>'.format(f=font)]
    if bold:
        rpr.append("<w:b/><w:bCs/>")
    if italic:
        rpr.append("<w:i/><w:iCs/>")
    if underline:
        rpr.append('<w:u w:val="single"/>')
    if color:
        rpr.append('<w:color w:val="{c}"/>'.format(c=color))
    rpr.append('<w:sz w:val="{s}"/><w:szCs w:val="{s}"/>'.format(s=size))
    rpr_xml = "<w:rPr>" + "".join(rpr) + "</w:rPr>"
    text_xml = ('<w:t xml:space="preserve">' + escape(text) + "</w:t>"
                if text else "<w:t/>")
    return "<w:r>" + rpr_xml + text_xml + "</w:r>"


def field(width=30, *, bold=False):
    """Линеечка для заполнения от руки — подчёркнутые пробелы."""
    return run(" " * width, underline=True, bold=bold)


def runseq(*r):
    return "".join(r)


def para(content, *, align=None, spacing_before=0,
         spacing_after=SPACE_AFTER, indent_left=None,
         keep_next=False, keep_lines=False, line=LINE_SPACING):
    ppr_parts = []
    if keep_next:
        ppr_parts.append("<w:keepNext/>")
    if keep_lines:
        ppr_parts.append("<w:keepLines/>")
    ppr_parts.append(
        '<w:spacing w:before="{b}" w:after="{a}" '
        'w:line="{l}" w:lineRule="auto"/>'.format(
            b=spacing_before, a=spacing_after, l=line))
    if align:
        ppr_parts.append('<w:jc w:val="{a}"/>'.format(a=align))
    if indent_left is not None:
        ppr_parts.append('<w:ind w:left="{i}"/>'.format(i=indent_left))
    ppr = "<w:pPr>" + "".join(ppr_parts) + "</w:pPr>" if ppr_parts else ""
    return "<w:p>" + ppr + content + "</w:p>"


def title(text):
    return para(
        run(text, bold=True, font=FONT_HEAD, size=SIZE_TITLE),
        align="center", spacing_after=120, keep_next=True,
    )


def h2(text):
    return para(
        run(text, bold=True, font=FONT_HEAD, size=SIZE_H2),
        spacing_before=80, spacing_after=SPACE_AFTER_HEAD,
        keep_next=True, keep_lines=True,
    )


def h3(text):
    return para(
        run(text, bold=True, font=FONT_TEXT, size=SIZE_H3),
        spacing_before=40, spacing_after=20,
        keep_next=True, keep_lines=True,
    )


def p(*parts, indent=None, align=None, spacing_after=SPACE_AFTER):
    return para("".join(parts), align=align, indent_left=indent,
                spacing_after=spacing_after)


def p_note(text):
    return para(
        run(text, italic=True, size=SIZE_NOTE, color="595959"),
        spacing_after=SPACE_AFTER,
    )


# --------------------------- table helpers ---------------------------


def _no_borders():
    return (
        "<w:tcBorders>"
        '<w:top w:val="nil"/><w:left w:val="nil"/>'
        '<w:bottom w:val="nil"/><w:right w:val="nil"/>'
        "</w:tcBorders>"
    )


def cell(width, content, *, valign="top"):
    return (
        "<w:tc>"
        '<w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{b}'
        '<w:vAlign w:val="{v}"/></w:tcPr>'
        "{c}"
        "</w:tc>"
    ).format(w=width, b=_no_borders(), v=valign, c=content)


def cell_span(width, content, span):
    return (
        "<w:tc>"
        '<w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{b}'
        '<w:gridSpan w:val="{s}"/></w:tcPr>'
        "{c}"
        "</w:tc>"
    ).format(w=width, b=_no_borders(), s=span, c=content)


def table(rows, col_widths):
    grid = "<w:tblGrid>" + "".join(
        '<w:gridCol w:w="{w}"/>'.format(w=w) for w in col_widths
    ) + "</w:tblGrid>"
    tbl_pr = (
        "<w:tblPr>"
        '<w:tblW w:w="{w}" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        "<w:tblCellMargin>"
        '<w:top w:w="20" w:type="dxa"/><w:left w:w="40" w:type="dxa"/>'
        '<w:bottom w:w="20" w:type="dxa"/><w:right w:w="40" w:type="dxa"/>'
        "</w:tblCellMargin>"
        "</w:tblPr>"
    ).format(w=sum(col_widths))
    return "<w:tbl>" + tbl_pr + grid + "".join(rows) + "</w:tbl>"


def tr(*cells):
    return "<w:tr>" + "".join(cells) + "</w:tr>"


# --------------------------- содержимое ---------------------------

LABEL_W = 2400
VALUE_W = 7800


def party_table_arendator():
    rows = []
    for label in [
        "ФИО:",
        "Зарегистрирован(а) по адресу:",
        "Паспорт серия и номер:",
        "Кем и когда выдан:",
        "Телефон:",
    ]:
        rows.append(
            tr(
                cell(LABEL_W, para(run(label), spacing_after=0)),
                cell(VALUE_W, para(field(80), spacing_after=0)),
            )
        )
    return table(rows, [LABEL_W, VALUE_W])


def options_table():
    rows = [
        tr(
            cell(400, para(run("☐"), spacing_after=0)),
            cell(4500, para(run("Насос электрический"), spacing_after=0)),
            cell(1500, para(runseq(field(8), run(" шт.")), spacing_after=0)),
            cell(3800, para(run("+ 400 ₽ за весь срок аренды", bold=True),
                            spacing_after=0)),
        ),
        tr(
            cell(400, para(run("☐"), spacing_after=0)),
            cell(4500, para(run("Спасательный жилет"), spacing_after=0)),
            cell(1500, para(runseq(field(8), run(" шт.")), spacing_after=0)),
            cell(3800, para(run("бесплатно", bold=True), spacing_after=0)),
        ),
        tr(
            cell(400, para(run("☐"), spacing_after=0)),
            cell(4500, para(run("Доставка Оборудования"), spacing_after=0)),
            cell(1500, para(run(""), spacing_after=0)),
            cell(3800, para(runseq(run("+ "), field(10),
                                   run(" ₽ (по согласованию)")),
                            spacing_after=0)),
        ),
    ]
    return table(rows, [400, 4500, 1500, 3800])


def _sig_row(label, fio):
    fio_text = run(fio, bold=True) if fio else field(20)
    return tr(
        cell(4500, para(run(label), spacing_after=0)),
        cell(3500, para(field(28), spacing_after=0)),
        cell(2200, para(runseq(run("/ "), fio_text, run(" /")),
                        spacing_after=0)),
    )


def _sig_subhead(text):
    inner = para(run(text, bold=True, size=SIZE_H3),
                 spacing_after=0, spacing_before=80)
    return "<w:tr>" + cell_span(10200, inner, 3) + "</w:tr>"


def signatures_table():
    rows = [
        _sig_row("Арендодатель сдал Оборудование:", "Воробьев П.А."),
        _sig_row("Арендатор принял Оборудование:", None),
        _sig_subhead("7.2. Возврат Оборудования (окончание срока):"),
        _sig_row("Арендатор сдал Оборудование:", None),
        _sig_row("Арендодатель принял Оборудование:", "Воробьев П.А."),
    ]
    return table(rows, [4500, 3500, 2200])


# --------------------------- сборка тела ---------------------------


def build_body():
    parts = []

    parts.append(title("ДОГОВОР ОКАЗАНИЯ УСЛУГ ПРОКАТА СПОРТИВНОГО ИНВЕНТАРЯ"))

    # дата + место
    parts.append(
        p(
            run("г. Рязань"),
            run("\t\t\t\t\t\t\t\t\t\t"),
            run("«"), field(3), run("» "), field(14), run(" 20"),
            field(2), run(" г."),
        )
    )

    # Арендодатель
    parts.append(
        p(
            run("Воробьев Павел Алексеевич", bold=True),
            run(", паспорт "),
            run("61 24 № 381913", bold=True),
            run(", выдан УМВД России по Рязанской области, "
                "зарегистрирован по адресу: г. Рязань, пос. Солотча, "
                "ул. Гайдара, д. 3, кв. 8, тел. "),
            run("8 915 605 38 03", bold=True),
            run(", именуемый в дальнейшем "),
            run("«Арендодатель»", bold=True),
            run(", с одной стороны,"),
        )
    )

    parts.append(
        p(
            run("и Гражданин(ка), указанный(ая) ниже, именуемый(ая) "
                "в дальнейшем "),
            run("«Арендатор»", bold=True),
            run(", с другой стороны, заключили настоящий Договор:"),
            spacing_after=20,
        )
    )

    parts.append(party_table_arendator())

    parts.append(p_note(
        "Личность Арендатора удостоверяется паспортом при заключении "
        "Договора. Подлинник паспорта в залог не изымается."
    ))

    # ----- 1. Предмет -----
    parts.append(h2("1. ПРЕДМЕТ ДОГОВОРА И КОМПЛЕКТАЦИЯ"))

    parts.append(
        p(
            run("1.1.", bold=True),
            run(" Арендодатель передаёт Арендатору во временное пользование "
                "комплект(ы) оборудования (далее — "),
            run("«Оборудование»", bold=True),
            run("). В состав одного "),
            run("базового комплекта", bold=True),
            run(" (входит в стоимость аренды) входит: а) SUP-board надувной "
                "Funwater; б) плавник пластиковый; в) насос ручной "
                "механический с манометром, клапаном и шлангом; "
                "г) весло сборное трёхсекционное; д) рюкзак для "
                "транспортировки."),
        )
    )

    parts.append(
        p(
            run("1.2.", bold=True),
            run(" Количество передаваемых базовых комплектов: "),
            field(6, bold=True),
            run(" шт."),
        )
    )

    parts.append(
        p(
            run("1.3.", bold=True),
            run(" Дополнительные опции (по желанию Арендатора, отметить ☑):"),
            spacing_after=20,
        )
    )
    parts.append(options_table())

    # ----- 2. Срок и расчёты -----
    parts.append(h2("2. СРОК АРЕНДЫ И РАСЧЁТЫ"))

    parts.append(
        p(
            run("2.1.", bold=True),
            run(" Срок аренды — "),
            field(5, bold=True),
            run(" сутки/суток с момента подписания настоящего Договора. "
                "По соглашению Сторон Оборудование может быть возвращено "
                "ранее окончания срока."),
        )
    )

    parts.append(
        p(
            run("2.2.", bold=True),
            run(" Арендная плата за весь срок пользования Оборудованием "
                "с учётом выбранных в п. 1.3 опций составляет "),
            field(12, bold=True),
            run(" ("), field(28), run(") рублей."),
        )
    )
    parts.append(
        p(
            run("2.3.", bold=True),
            run(" Залог не взимается. Арендодатель сверяет данные паспорта "
                "Арендатора при передаче Оборудования."),
        )
    )

    # ----- 3. Обязанности -----
    parts.append(h2("3. ОБЯЗАННОСТИ СТОРОН"))

    parts.append(h3("3.1. Арендодатель обязан:"))
    parts.append(p(run("3.1.1.", bold=True),
                   run(" Передать Арендатору Оборудование в рабочем "
                       "состоянии в срок, указанный в п. 2.1."), indent=300))
    parts.append(p(run("3.1.2.", bold=True),
                   run(" Оказывать консультативную помощь по использованию "
                       "Оборудования."), indent=300))
    parts.append(p(run("3.1.3.", bold=True),
                   run(" Принять Оборудование по окончании срока аренды."),
                   indent=300))

    parts.append(h3("3.2. Арендатор обязан:"))
    parts.append(p(
        run("3.2.1.", bold=True),
        run(" Использовать Оборудование в соответствии с его назначением. "
            "При нарушении назначения или условий Договора Арендодатель "
            "вправе потребовать расторжения Договора и возмещения убытков."),
        indent=300,
    ))
    parts.append(p(run("3.2.2.", bold=True),
                   run(" Поддерживать Оборудование в исправном состоянии."),
                   indent=300))
    parts.append(p(
        run("3.2.3.", bold=True),
        run(" Возвратить Оборудование в том же состоянии, в каком оно было "
            "передано (с учётом нормального износа), и в "),
        run("чистом виде", bold=True),
        run(" — без песка, грязи, травы, земли и иных явных загрязнений — "
            "не позднее срока, указанного в п. 2.1. "),
        run("Штраф за возврат загрязнённого Оборудования — 500 (пятьсот) "
            "рублей.", bold=True),
        indent=300,
    ))

    # ----- 4. Ответственность -----
    parts.append(h2("4. ОТВЕТСТВЕННОСТЬ СТОРОН"))

    parts.append(p(run("4.1.", bold=True),
                   run(" Стороны несут ответственность за неисполнение или "
                       "ненадлежащее исполнение условий Договора.")))
    parts.append(p(
        run("4.2.", bold=True),
        run(" Арендодатель не отвечает за недостатки Оборудования, "
            "оговорённые при заключении Договора, заранее известные "
            "Арендатору либо обнаруживаемые при обычном осмотре при выдаче."),
    ))
    parts.append(p(
        run("4.3.", bold=True),
        run(" При повреждении или утрате Оборудования по вине Арендатора "
            "(или третьих лиц, за которых он отвечает) Арендатор возмещает "
            "Арендодателю стоимость согласно следующему прайсу."),
    ))
    parts.append(p(
        run("4.3.1.", bold=True),
        run(" Полная утрата или повреждение без возможности ремонта: "),
        run("SUP-board надувной Funwater — 20 000 ₽", bold=True),
        run("; "),
        run("плавник пластиковый — 2 000 ₽", bold=True),
        run("; "),
        run("насос ручной механический (с манометром, клапаном, шлангом) "
            "— 3 000 ₽", bold=True),
        run("; "),
        run("весло сборное трёхсекционное — 3 500 ₽", bold=True),
        run("; "),
        run("рюкзак для транспортировки — 2 500 ₽", bold=True),
        run("; "),
        run("насос электрический — 3 500 ₽", bold=True),
        run("; "),
        run("спасательный жилет — 1 500 ₽", bold=True),
        run("."),
        indent=300,
    ))
    parts.append(p(
        run("4.3.2.", bold=True),
        run(" Повреждение или утрата с возможностью ремонта/замены: "
            "стоимость ремонта/замены части плюс "),
        run("2 500 (две тысячи пятьсот) рублей", bold=True),
        run(" независимо от сложности восстановления."),
        indent=300,
    ))
    parts.append(p(
        run("4.4.", bold=True),
        run(" Арендатор не отвечает за невозможность Арендодателя принять "
            "Оборудование в обозначенный срок."),
    ))

    # ----- 5. Споры -----
    parts.append(h2("5. ПОРЯДОК РАЗРЕШЕНИЯ СПОРОВ"))
    parts.append(p(
        run("5.1.", bold=True),
        run(" Все споры разрешаются переговорами. При невозможности — в "
            "суде по месту нахождения Арендодателя в порядке, установленном "
            "законодательством РФ."),
    ))

    # ----- 6. Заключительные -----
    parts.append(h2("6. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ"))
    parts.append(p(
        run("6.1.", bold=True),
        run(" Договор вступает в силу с момента подписания Сторонами. "
            "Окончание срока действия Договора не освобождает Стороны от "
            "обязательств, возникших до его окончания."),
    ))
    parts.append(p(
        run("6.2.", bold=True),
        run(" Изменения и дополнения оформляются письменно и подписываются "
            "обеими Сторонами; равноценны изменения, зафиксированные в "
            "электронной переписке Сторон."),
    ))
    parts.append(p(
        run("6.3.", bold=True),
        run(" Договор составляется в одном экземпляре, который остаётся у "
            "Арендодателя; по требованию Арендатора оформляется второй "
            "экземпляр, либо предоставляется фотокопия. Во всём остальном "
            "Стороны руководствуются законодательством РФ."),
    ))

    # ----- 7. Подписи / акт -----
    parts.append(h2("7. ПОДПИСИ СТОРОН (АКТ ПРИЁМА-ПЕРЕДАЧИ)"))
    parts.append(p_note(
        "Реквизиты Сторон указаны в шапке настоящего Договора и не "
        "повторяются. 7.1 — выдача, 7.2 — возврат."
    ))
    parts.append(h3("7.1. Выдача Оборудования (начало срока аренды):"))
    parts.append(signatures_table())

    return "".join(parts)


# --------------------------- сборка docx ---------------------------


def docx_files():
    body = build_body()

    sect_pr = (
        "<w:sectPr>"
        '<w:pgSz w:w="11906" w:h="16838"/>'
        # узкие поля: 1.4cm слева/справа, 1.0cm сверху/снизу
        '<w:pgMar w:top="567" w:right="794" w:bottom="567" w:left="794" '
        'w:header="284" w:footer="284" w:gutter="0"/>'
        "</w:sectPr>"
    )

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document ' + W + '>'
        "<w:body>" + body + sect_pr + "</w:body>"
        "</w:document>"
    )

    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles ' + W + '>'
        "<w:docDefaults>"
        "<w:rPrDefault><w:rPr>"
        '<w:rFonts w:ascii="' + FONT_TEXT + '" w:hAnsi="' + FONT_TEXT + '" '
        'w:cs="' + FONT_TEXT + '"/>'
        '<w:sz w:val="' + str(SIZE_TEXT) + '"/>'
        '<w:szCs w:val="' + str(SIZE_TEXT) + '"/>'
        '<w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>'
        "</w:rPr></w:rPrDefault>"
        "<w:pPrDefault><w:pPr>"
        '<w:spacing w:before="0" w:after="' + str(SPACE_AFTER) + '" '
        'w:line="' + str(LINE_SPACING) + '" w:lineRule="auto"/>'
        '<w:jc w:val="both"/>'
        "</w:pPr></w:pPrDefault>"
        "</w:docDefaults>"
        "</w:styles>"
    )

    settings = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings ' + W + '>'
        '<w:zoom w:percent="100"/>'
        '<w:defaultTabStop w:val="708"/>'
        "</w:settings>"
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
        'content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-'
        'package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        "</Types>"
    )

    rels_root = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )

    rels_doc = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        'relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/settings" Target="settings.xml"/>'
        "</Relationships>"
    )

    return {
        "[Content_Types].xml": content_types.encode("utf-8"),
        "_rels/.rels": rels_root.encode("utf-8"),
        "word/_rels/document.xml.rels": rels_doc.encode("utf-8"),
        "word/document.xml": document.encode("utf-8"),
        "word/styles.xml": styles.encode("utf-8"),
        "word/settings.xml": settings.encode("utf-8"),
    }


def main():
    here = Path(__file__).resolve().parent.parent
    out = here / "docs" / "contract_v2.docx"
    out.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in docx_files().items():
            zf.writestr(name, data)

    print("Wrote", out, "size", os.path.getsize(out), "bytes")


if __name__ == "__main__":
    main()
