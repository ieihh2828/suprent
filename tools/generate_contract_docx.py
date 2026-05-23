#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор docs/contract_v2.docx — Word-версия договора проката
из шаблона docs/contract_v2_template.md (по содержанию идентичны).

Без внешних зависимостей: docx-файл собирается через стандартный
zipfile + ручная сборка Office Open XML (минимально достаточная схема,
открывается Word/LibreOffice/Google Docs).
"""

from __future__ import annotations

import os
import zipfile
from html import escape
from pathlib import Path

# --------------------------- параметры стилей ---------------------------

FONT_TEXT = "Times New Roman"
FONT_HEAD = "Arial"

# размеры в "half-points" — Word считает в полупунктах (16pt = 32)
SIZE_TITLE = 32     # 16 pt
SIZE_H2 = 24        # 12 pt
SIZE_H3 = 22        # 11 pt
SIZE_TEXT = 22      # 11 pt
SIZE_NOTE = 20      # 10 pt

# --------------------------- helpers ---------------------------

W = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""

def run(text: str, *, bold=False, italic=False, font=FONT_TEXT,
        size=SIZE_TEXT, color: str | None = None,
        underline=False) -> str:
    """Один <w:r> (run) с настройками шрифта."""
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

    text_xml = (
        '<w:t xml:space="preserve">' + escape(text) + "</w:t>"
        if text else "<w:t/>"
    )
    return "<w:r>" + rpr_xml + text_xml + "</w:r>"


def runs(*parts) -> str:
    return "".join(parts)


def para(content: str, *, align: str | None = None,
         spacing_before: int = 0, spacing_after: int = 120,
         indent_left: int | None = None,
         border: bool = False, shading: str | None = None,
         keep_next: bool = False) -> str:
    """Один абзац <w:p> с форматированием."""
    ppr_parts = []
    if keep_next:
        ppr_parts.append("<w:keepNext/>")
    ppr_parts.append(
        '<w:spacing w:before="{b}" w:after="{a}"/>'.format(
            b=spacing_before, a=spacing_after))
    if align:
        ppr_parts.append('<w:jc w:val="{a}"/>'.format(a=align))
    if indent_left is not None:
        ppr_parts.append('<w:ind w:left="{i}"/>'.format(i=indent_left))
    if border:
        ppr_parts.append(
            "<w:pBdr>"
            '<w:top w:val="single" w:sz="18" w:space="6" w:color="C00000"/>'
            '<w:left w:val="single" w:sz="18" w:space="6" w:color="C00000"/>'
            '<w:bottom w:val="single" w:sz="18" w:space="6" w:color="C00000"/>'
            '<w:right w:val="single" w:sz="18" w:space="6" w:color="C00000"/>'
            "</w:pBdr>"
        )
    if shading:
        ppr_parts.append(
            '<w:shd w:val="clear" w:color="auto" w:fill="{s}"/>'.format(
                s=shading))
    ppr = "<w:pPr>" + "".join(ppr_parts) + "</w:pPr>" if ppr_parts else ""
    return "<w:p>" + ppr + content + "</w:p>"


def title(text: str) -> str:
    return para(
        run(text, bold=True, font=FONT_HEAD, size=SIZE_TITLE),
        align="center", spacing_before=0, spacing_after=240,
    )


def h2(text: str) -> str:
    return para(
        run(text, bold=True, font=FONT_HEAD, size=SIZE_H2),
        spacing_before=240, spacing_after=120, keep_next=True,
    )


def h3(text: str) -> str:
    return para(
        run(text, bold=True, font=FONT_TEXT, size=SIZE_H3),
        spacing_before=120, spacing_after=60, keep_next=True,
    )


def p_text(*parts: str, align: str | None = None,
           indent: int | None = None) -> str:
    return para("".join(parts), align=align, indent_left=indent)


def p_note(text: str) -> str:
    return para(
        run(text, italic=True, size=SIZE_NOTE, color="595959"),
        spacing_after=120,
    )


def underscore(n: int) -> str:
    """Подчерк для заполнения от руки (визуально пробел с подчёркиванием)."""
    return "_" * n


def sep_line() -> str:
    """Тонкая горизонтальная линия (граница абзаца снизу)."""
    pPr = (
        "<w:pPr>"
        '<w:spacing w:before="0" w:after="0"/>'
        "<w:pBdr>"
        '<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>'
        "</w:pBdr>"
        "</w:pPr>"
    )
    return "<w:p>" + pPr + "</w:p>"


def warning_box(title_text: str, body_lines: list[str]) -> str:
    """Красная плашка-предупреждение: рамка + светло-красный фон."""
    out = []
    # заголовок плашки — крупнее, красный, жирный
    out.append(
        para(
            run(title_text, bold=True, font=FONT_HEAD,
                size=SIZE_H2, color="C00000"),
            border=True, shading="FBE4E4",
            spacing_before=240, spacing_after=0,
            keep_next=True,
        )
    )
    # пустая граничная строка для слитного блока
    for i, line in enumerate(body_lines):
        is_last = i == len(body_lines) - 1
        out.append(
            para(
                run(line, size=SIZE_TEXT, color="C00000"),
                border=True, shading="FBE4E4",
                spacing_before=0,
                spacing_after=120 if is_last else 0,
            )
        )
    return "".join(out)


# --------------------------- содержимое ---------------------------

def build_body() -> str:
    parts: list[str] = []

    parts.append(title("ДОГОВОР ОКАЗАНИЯ УСЛУГ ПРОКАТА СПОРТИВНОГО ИНВЕНТАРЯ"))

    # дата + место заключения
    parts.append(
        p_text(
            run("г. Рязань", size=SIZE_TEXT),
            run("\t\t\t\t\t\t\t\t", size=SIZE_TEXT),
            run("«", size=SIZE_TEXT),
            run(underscore(3), size=SIZE_TEXT),
            run("» ", size=SIZE_TEXT),
            run(underscore(14), size=SIZE_TEXT),
            run(" 20", size=SIZE_TEXT),
            run(underscore(2), size=SIZE_TEXT),
            run(" г.", size=SIZE_TEXT),
        )
    )

    # Арендодатель — предзаполнен
    parts.append(
        p_text(
            run("Воробьев Павел Алексеевич", bold=True),
            run(", паспорт серия и номер "),
            run("61 24 № 381913", bold=True),
            run(", выдан "),
            run("УМВД России по Рязанской области", bold=True),
            run(", зарегистрирован по адресу: "),
            run("г. Рязань, пос. Солотча, ул. Гайдара, д. 3, кв. 8", bold=True),
            run(", телефон: "),
            run("89156053803", bold=True),
            run(", именуемый в дальнейшем "),
            run("«Арендодатель»", bold=True),
            run(", с одной стороны,"),
        )
    )

    # Арендатор — поля под заполнение
    parts.append(
        p_text(
            run("и Гражданин(ка) "),
            run(underscore(80)),
            run(","),
        )
    )
    parts.append(
        p_text(
            run("зарегистрированный(ая) по адресу: "),
            run(underscore(70)),
            run(","),
        )
    )
    parts.append(
        p_text(
            run("паспорт серия и номер "),
            run(underscore(75)),
            run(","),
        )
    )
    parts.append(
        p_text(
            run("выдан "),
            run(underscore(85)),
            run(","),
        )
    )
    parts.append(
        p_text(
            run("телефон: "),
            run(underscore(40)),
            run(", именуемый(ая) в дальнейшем "),
            run("«Арендатор»", bold=True),
            run(", с другой стороны,"),
        )
    )
    parts.append(
        p_text(
            run("именуемые вместе — "),
            run("«Стороны»", bold=True),
            run(", заключили настоящий Договор о нижеследующем."),
        )
    )

    parts.append(
        p_note(
            "Личность Арендатора удостоверяется паспортом при заключении "
            "настоящего Договора. Подлинник паспорта в залог не изымается."
        )
    )

    # ----- 1. Предмет -----
    parts.append(h2("1. ПРЕДМЕТ ДОГОВОРА И КОМПЛЕКТАЦИЯ"))

    parts.append(
        p_text(
            run("1.1.", bold=True),
            run(" Арендодатель передаёт Арендатору на возмездной основе "
                "во временное пользование исправный(ые) и укомплектованный(ые) "
                "комплект(ы) оборудования (далее — "),
            run("«Оборудование»", bold=True),
            run("), а Арендатор принимает их."),
        )
    )

    parts.append(
        p_text(
            run("1.2.", bold=True),
            run(" В состав одного "),
            run("базового комплекта", bold=True),
            run(" (входит в стоимость аренды) входит:"),
        )
    )
    for line in [
        "а) SUP-board надувной Funwater;",
        "б) плавник пластиковый;",
        "в) насос ручной механический с манометром, клапаном и шлангом;",
        "г) весло сборное трёхсекционное (металл/пластик);",
        "д) рюкзак для транспортировки Оборудования.",
    ]:
        parts.append(p_text(run(line), indent=420))

    parts.append(
        p_text(
            run("1.3.", bold=True),
            run(" Количество передаваемых базовых комплектов: "),
            run(underscore(6), bold=True),
            run(" комплект(а/ов)."),
        )
    )

    parts.append(
        p_text(
            run("1.4.", bold=True),
            run(" Дополнительные опции (по желанию Арендатора, "
                "отметить знаком ✓ или обвести кружком):"),
        )
    )
    for ch_line in [
        ("☐ Насос электрический",
         underscore(5) + " шт.",
         "+ 400 ₽ за весь срок аренды"),
        ("☐ Спасательный жилет",
         underscore(5) + " шт.",
         "бесплатно"),
        ("☐ Доставка Оборудования",
         "",
         "+ " + underscore(7) + " ₽ (по согласованию)"),
    ]:
        left, qty, price = ch_line
        parts.append(
            p_text(
                run(left + "    "),
                run(qty + "    " if qty else ""),
                run(price, bold=True),
                indent=420,
            )
        )

    # ----- 2. Срок аренды и расчёты -----
    parts.append(h2("2. СРОК АРЕНДЫ И РАСЧЁТЫ"))

    parts.append(
        p_text(
            run("2.1.", bold=True),
            run(" Срок аренды (отметить применимый вариант):"),
        )
    )
    parts.append(
        p_text(
            run("☐ "),
            run("Почасовая аренда:", bold=True),
            run(" с "),
            run(underscore(5), bold=True),
            run(" : "),
            run(underscore(5), bold=True),
            run(" до "),
            run(underscore(5), bold=True),
            run(" : "),
            run(underscore(5), bold=True),
            run(", итого "),
            run(underscore(5), bold=True),
            run(" час(а/ов)."),
            indent=420,
        )
    )
    parts.append(
        p_text(
            run("☐ "),
            run("Аренда на сутки:", bold=True),
            run(" "),
            run(underscore(5), bold=True),
            run(" суток, возврат не позднее "),
            run(underscore(5), bold=True),
            run(" : "),
            run(underscore(5), bold=True),
            run(" последних суток аренды."),
            indent=420,
        )
    )

    parts.append(
        p_text(
            run("2.2.", bold=True),
            run(" Арендная плата по настоящему Договору составляет "),
            run(underscore(12), bold=True),
            run(" ("),
            run(underscore(28), bold=True),
            run(") рублей за весь срок пользования Оборудованием с учётом "
                "выбранных в п. 1.4 дополнительных опций."),
        )
    )

    parts.append(
        p_text(
            run("2.3.", bold=True),
            run(" Залог по настоящему Договору не взимается. В качестве "
                "идентификации Арендатора Арендодатель сверяет данные паспорта "
                "Арендатора при заключении настоящего Договора и передаче "
                "Оборудования."),
        )
    )

    parts.append(
        p_text(
            run("2.4.", bold=True),
            run(" Подписанием настоящего Договора Стороны подтверждают:"),
        )
    )
    parts.append(
        p_text(
            run("— Арендодатель — получение от Арендатора денежных средств за "
                "аренду Оборудования;"),
            indent=420,
        )
    )
    parts.append(
        p_text(
            run("— Арендатор — получение от Арендодателя Оборудования, "
                "указанного в пп. 1.2–1.4 настоящего Договора."),
            indent=420,
        )
    )

    # ----- 3. Обязанности -----
    parts.append(h2("3. ОБЯЗАННОСТИ СТОРОН"))

    parts.append(h3("3.1. Арендодатель обязан:"))
    parts.append(
        p_text(
            run("3.1.1.", bold=True),
            run(" Передать Арендатору Оборудование в рабочем состоянии, "
                "отвечающем условиям настоящего Договора, в срок, указанный в "
                "п. 2.1."),
        )
    )
    parts.append(
        p_text(
            run("3.1.2.", bold=True),
            run(" Оказывать консультативную и иную помощь в рамках "
                "использования Оборудования."),
        )
    )
    parts.append(
        p_text(
            run("3.1.3.", bold=True),
            run(" Принять Оборудование по окончании срока аренды."),
        )
    )

    parts.append(h3("3.2. Арендатор обязан:"))
    parts.append(
        p_text(
            run("3.2.1.", bold=True),
            run(" Использовать Оборудование в соответствии с условиями "
                "Договора и назначением имущества. Если Арендатор пользуется "
                "Оборудованием с нарушениями условий Договора или назначения "
                "имущества, Арендодатель имеет право потребовать расторжения "
                "Договора и возмещения убытков."),
        )
    )
    parts.append(
        p_text(
            run("3.2.2.", bold=True),
            run(" Поддерживать Оборудование в исправном состоянии."),
        )
    )
    parts.append(
        p_text(
            run("3.2.3.", bold=True),
            run(" Возвратить Оборудование Арендодателю в том же состоянии, "
                "в каком оно было передано, с учётом нормального износа, не "
                "позднее срока, указанного в п. 2.1 настоящего Договора."),
        )
    )

    # плашка штрафа
    parts.append(
        warning_box(
            "⚠ ВНИМАНИЕ: ШТРАФ ЗА ЗАГРЯЗНЁННОЕ ОБОРУДОВАНИЕ — 700 ₽",
            [
                "Оборудование возвращается в чистом виде: без песка, грязи, "
                "травы, земли и иных сильно бросающихся в глаза загрязнений.",
                "При возврате Оборудования с любым из перечисленных загрязнений "
                "Арендатор оплачивает штраф 700 (семьсот) рублей дополнительно "
                "к арендной плате (п. 3.2.4).",
            ],
        )
    )

    parts.append(
        p_text(
            run("3.2.4.", bold=True),
            run(" Возвратить Оборудование в чистом виде. Штраф за возврат "
                "загрязнённого Оборудования — "),
            run("700 (семьсот) рублей", bold=True),
            run("."),
        )
    )

    # ----- 4. Ответственность -----
    parts.append(h2("4. ОТВЕТСТВЕННОСТЬ СТОРОН"))

    parts.append(
        p_text(
            run("4.1.", bold=True),
            run(" Стороны несут ответственность за неисполнение или "
                "ненадлежащее исполнение условий Договора."),
        )
    )
    parts.append(
        p_text(
            run("4.2.", bold=True),
            run(" Арендодатель не отвечает за недостатки сданного в аренду "
                "Оборудования, которые были им оговорены при заключении "
                "Договора аренды или были заранее известны Арендатору либо "
                "должны были быть обнаружены Арендатором во время осмотра "
                "Оборудования или проверки его исправности при заключении "
                "Договора и/или передаче Оборудования в аренду."),
        )
    )
    parts.append(
        p_text(
            run("4.3.", bold=True),
            run(" При возврате неисправного арендованного Оборудования, "
                "поврежденного по вине Арендатора или третьих лиц, "
                "ответственность за которых в отношении Оборудования несёт "
                "Арендатор, Арендатор обязан возместить Арендодателю расходы "
                "по ремонту и/или утрате Оборудования:"),
        )
    )
    parts.append(
        p_text(
            run("4.3.1.", bold=True),
            run(" Полная утеря, утеря части или повреждение без возможности "
                "ремонта Оборудования:"),
        )
    )
    for line in [
        ("а) SUP-board надувной Funwater — ", "20 000", " рублей;"),
        ("б) плавник пластиковый — ", "2 000", " рублей;"),
        ("в) насос ручной механический с манометром, клапаном и шлангом — ",
         "3 000", " рублей;"),
        ("г) весло сборное трёхсекционное — ", "3 500", " рублей;"),
        ("д) рюкзак для транспортировки Оборудования — ", "2 500",
         " рублей;"),
        ("е) насос электрический — ", "3 500", " рублей;"),
        ("ж) спасательный жилет — ", "1 500", " рублей."),
    ]:
        prefix, amount, suffix = line
        parts.append(
            p_text(
                run(prefix),
                run(amount, bold=True),
                run(suffix),
                indent=420,
            )
        )
    parts.append(
        p_text(
            run("4.3.2.", bold=True),
            run(" Повреждение или утеря части Оборудования с возможностью "
                "ремонта или частичной замены рассчитывается по формуле: "
                "стоимость ремонта/возмещения/покупки части Оборудования плюс "),
            run("2 500 (две тысячи пятьсот) рублей", bold=True),
            run(" независимо от сложности восстановления Оборудования к "
                "рабочему/исправному/пригодному к дальнейшей эксплуатации "
                "состоянию."),
        )
    )
    parts.append(
        p_text(
            run("4.4.", bold=True),
            run(" Арендатор не несёт ответственность за невозможность "
                "Арендодателя принять Оборудование в обозначенные сроки по "
                "Договору."),
        )
    )

    # ----- 5. Споры -----
    parts.append(h2("5. ПОРЯДОК РАЗРЕШЕНИЯ СПОРОВ"))
    parts.append(
        p_text(
            run("5.1.", bold=True),
            run(" Все споры или разногласия, возникающие между Сторонами по "
                "настоящему Договору или возникшие в связи с ним, разрешаются "
                "путём переговоров между Сторонами."),
        )
    )
    parts.append(
        p_text(
            run("5.2.", bold=True),
            run(" В случае невозможности разрешения разногласий путём "
                "переговоров они подлежат рассмотрению в суде в порядке, "
                "установленном действующим законодательством Российской "
                "Федерации."),
        )
    )

    # ----- 6. Заключительные -----
    parts.append(h2("6. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ"))
    for n, text in [
        ("6.1.", "Настоящий Договор вступает в силу с момента его подписания "
                 "Сторонами."),
        ("6.2.", "Любые изменения и дополнения к настоящему Договору имеют "
                 "силу только в том случае, если они оформлены в письменном "
                 "виде и подтверждены обеими Сторонами. Принимаются изменения, "
                 "зафиксированные в ходе электронной переписки Сторон, в "
                 "случае их фиксации обеими Сторонами."),
        ("6.3.", "Окончание срока действия Договора не освобождает Стороны от "
                 "исполнения обязательств по Договору."),
        ("6.4.", "Настоящий Договор может быть расторгнут досрочно по "
                 "письменному соглашению Сторон или в иных случаях, "
                 "предусмотренных законом Российской Федерации."),
        ("6.5.", "Настоящий Договор составляется в одном экземпляре, который "
                 "остаётся у Арендодателя. По письменному или устному "
                 "требованию Арендатора может быть оформлен второй экземпляр "
                 "на бумажном носителе. В отсутствие такого требования "
                 "Арендатор вправе получить фотокопию (скан-копию) подписанного "
                 "Договора, имеющую информационное значение."),
        ("6.6.", "Во всём остальном Стороны руководствуются законодательством "
                 "Российской Федерации."),
    ]:
        parts.append(
            p_text(
                run(n, bold=True),
                run(" " + text),
            )
        )

    # ----- 7. Подписи -----
    parts.append(h2("7. ПОДПИСИ СТОРОН (АКТ ПРИЁМА-ПЕРЕДАЧИ)"))
    parts.append(
        p_note("Реквизиты Сторон указаны в шапке настоящего Договора и не "
               "повторяются.")
    )

    parts.append(h3("7.1. Выдача Оборудования (начало срока аренды):"))
    parts.append(
        p_text(
            run("Арендодатель сдал Оборудование:    "),
            run(underscore(28)),
            run("  /  "),
            run("Воробьев П.А.", bold=True),
            run("  /"),
        )
    )
    parts.append(
        p_text(
            run("Арендатор принял Оборудование:    "),
            run(underscore(28)),
            run("  /  "),
            run(underscore(20)),
            run("  /"),
        )
    )

    parts.append(h3("7.2. Возврат Оборудования (окончание срока аренды):"))
    parts.append(
        p_text(
            run("Арендатор сдал Оборудование:        "),
            run(underscore(28)),
            run("  /  "),
            run(underscore(20)),
            run("  /"),
        )
    )
    parts.append(
        p_text(
            run("Арендодатель принял Оборудование:  "),
            run(underscore(28)),
            run("  /  "),
            run("Воробьев П.А.", bold=True),
            run("  /"),
        )
    )

    return "".join(parts)


# --------------------------- сборка docx ---------------------------

def docx_xml() -> dict[str, str | bytes]:
    body = build_body()

    # настройки страницы — A4 и поля
    sect_pr = (
        "<w:sectPr>"
        '<w:pgSz w:w="11906" w:h="16838"/>'  # A4 portrait
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
        'w:header="708" w:footer="708" w:gutter="0"/>'
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
        "<w:rPrDefault>"
        "<w:rPr>"
        '<w:rFonts w:ascii="' + FONT_TEXT + '" w:hAnsi="' + FONT_TEXT + '" '
        'w:cs="' + FONT_TEXT + '"/>'
        '<w:sz w:val="' + str(SIZE_TEXT) + '"/>'
        '<w:szCs w:val="' + str(SIZE_TEXT) + '"/>'
        '<w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/>'
        "</w:rPr>"
        "</w:rPrDefault>"
        "<w:pPrDefault>"
        "<w:pPr>"
        '<w:spacing w:before="0" w:after="120" w:line="276" '
        'w:lineRule="auto"/>'
        '<w:jc w:val="both"/>'
        "</w:pPr>"
        "</w:pPrDefault>"
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
        "[Content_Types].xml": content_types,
        "_rels/.rels": rels_root,
        "word/_rels/document.xml.rels": rels_doc,
        "word/document.xml": document,
        "word/styles.xml": styles,
        "word/settings.xml": settings,
    }


def main() -> None:
    here = Path(__file__).resolve().parent.parent
    out = here / "docs" / "contract_v2.docx"
    out.parent.mkdir(parents=True, exist_ok=True)

    files = docx_xml()

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            if isinstance(data, str):
                data = data.encode("utf-8")
            zf.writestr(name, data)

    print("Wrote", out, "size", os.path.getsize(out), "bytes")


if __name__ == "__main__":
    main()
