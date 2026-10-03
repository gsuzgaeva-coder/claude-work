#!/usr/bin/env python3
"""Сборка УПД (формат 5.03) с кодами маркировки для ЭДО Лайт «Честного знака».

Берёт исходный УПД (XML, windows-1251) и выгрузку КИЗ (xlsx: Заказ | Наименование товара | Киз),
раскладывает коды по строкам УПД по артикулу (КодТов = часть наименования до « | »),
проверяет количество и заменяет блок «Подписант» на сотрудника, действующего по доверенности.

Пример:
  python3 build_upd.py --upd input/upd.xml --kiz input/kiz.xlsx --out out \\
      --fio "Иванова Мария Петровна" --dolzhn "Менеджер" \\
      --mchd-nom 1b2c3d4e-... --mchd-date 01.09.2026
"""
import argparse
import collections
import re
import sys
import uuid
from pathlib import Path

import openpyxl
from lxml import etree

GS = "\x1d"


def kiz_without_tail(raw: str) -> str:
    """01+GTIN(14)+21+серийный номер — код идентификации без криптохвоста (91/92)."""
    code = raw.replace("_x001D_", GS).strip()
    ki = code.split(GS)[0]
    if not (ki.startswith("01") and ki[16:18] == "21" and len(ki) == 31):
        raise ValueError(f"Неожиданная структура КИЗ: {raw!r}")
    return ki


def load_kiz(path: Path) -> dict:
    ws = openpyxl.load_workbook(path, read_only=True).active
    by_art = collections.defaultdict(list)
    seen = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[2]:
            continue
        ki = kiz_without_tail(str(row[2]))
        if ki in seen:
            raise ValueError(f"Дубль КИЗ: {ki}")
        seen.add(ki)
        by_art[str(row[1]).split(" | ")[0].strip()].append(ki)
    return by_art


def fio_elem(fio: str) -> etree._Element:
    parts = fio.split()
    if len(parts) < 2:
        raise ValueError("ФИО: укажите как минимум фамилию и имя")
    el = etree.Element("ФИО", Фамилия=parts[0], Имя=parts[1])
    if len(parts) > 2:
        el.set("Отчество", " ".join(parts[2:]))
    return el


def build_signer(a) -> etree._Element:
    # ТипПодпис 1 — усиленная квалифицированная ЭП.
    # СпосПодтПолном 2 — по данным электронной доверенности (МЧД), 3 — по доверенности в бумажной форме.
    s = etree.Element("Подписант", Должн=a.dolzhn, ТипПодпис="1")
    if a.paper_nom:
        s.set("СпосПодтПолном", "3")
        s.append(fio_elem(a.fio))
        bum = etree.SubElement(s, "СвДоверБум", ДатаВыдДовер=a.paper_date, ВнНомДовер=a.paper_nom)
        bum.append(fio_elem(a.principal_fio))
    else:
        s.set("СпосПодтПолном", "2")
        s.append(fio_elem(a.fio))
        if a.mchd_nom:
            etree.SubElement(s, "СвДоверЭл", НомДовер=a.mchd_nom, ДатаВыдДовер=a.mchd_date,
                             ИдСистХран=a.mchd_system)
    return s


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--upd", required=True, type=Path)
    p.add_argument("--kiz", required=True, type=Path)
    p.add_argument("--out", required=True, type=Path)
    p.add_argument("--fio", required=True, help="ФИО сотрудника-подписанта")
    p.add_argument("--dolzhn", required=True, help="Должность сотрудника")
    p.add_argument("--mchd-nom", help="Номер (GUID) машиночитаемой доверенности")
    p.add_argument("--mchd-date", help="Дата выдачи МЧД, ДД.ММ.ГГГГ")
    p.add_argument("--mchd-system", default="https://m4d.nalog.gov.ru/",
                   help="Идентификатор системы хранения МЧД")
    p.add_argument("--paper-nom", help="Номер бумажной доверенности (вместо МЧД)")
    p.add_argument("--paper-date", help="Дата бумажной доверенности, ДД.ММ.ГГГГ")
    p.add_argument("--principal-fio", help="ФИО выдавшего бумажную доверенность")
    p.add_argument("--new-guid", action="store_true", help="Новый GUID в ИдФайл")
    a = p.parse_args()
    if a.mchd_nom and not a.mchd_date:
        p.error("--mchd-date обязателен вместе с --mchd-nom")
    if a.paper_nom and not (a.paper_date and a.principal_fio):
        p.error("для бумажной доверенности нужны --paper-date и --principal-fio")

    kiz = load_kiz(a.kiz)
    tree = etree.parse(str(a.upd), etree.XMLParser(remove_blank_text=True))
    root = tree.getroot()

    errors, used = [], set()
    total = 0
    for sv in root.iter("СведТов"):
        dop = sv.find("ДопСведТов")
        art = dop.get("КодТов")
        qty = int(float(sv.get("КолТов")))
        codes = kiz.get(art, [])
        if len(codes) != qty:
            errors.append(f"стр. {sv.get('НомСтр')} {art}: в УПД {qty}, КИЗ {len(codes)}")
            continue
        for old in dop.findall("НомСредИдентТов"):
            dop.remove(old)
        nsi = etree.SubElement(dop, "НомСредИдентТов")
        for c in codes:
            etree.SubElement(nsi, "КИЗ").text = c
        used.add(art)
        total += qty
    arts = {sv.find("ДопСведТов").get("КодТов") for sv in root.iter("СведТов")}
    errors += [f"КИЗ без строки в УПД: {art} ({len(kiz[art])} шт.)" for art in sorted(set(kiz) - arts)]
    if errors:
        print("ОШИБКИ сопоставления:\n  " + "\n  ".join(errors), file=sys.stderr)
        return 1

    doc = root.find("Документ")
    old = doc.find("Подписант")
    doc.replace(old, build_signer(a))

    guid_re = r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}"
    file_id = root.get("ИдФайл")
    if a.new_guid:
        file_id = re.sub(guid_re, str(uuid.uuid4()), file_id, count=1)
    # Второй признак после GUID = 1: в файле есть маркированные товары (как в прошлых УПД).
    file_id = re.sub(rf"({guid_re}_\d)_\d_", r"\1_1_", file_id, count=1)
    root.set("ИдФайл", file_id)

    a.out.mkdir(parents=True, exist_ok=True)
    out = a.out / f"{root.get('ИдФайл')}.xml"
    tree.write(str(out), encoding="windows-1251", xml_declaration=True, pretty_print=True)
    print(f"OK: {sum(1 for _ in root.iter('СведТов'))} строк, {total} КИЗ -> {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
