"""Извлечение и применение текста в PPTX без потери структуры пакета."""

from __future__ import annotations

import copy
import os
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Optional

from lxml import etree

from pptx_geometry import NS, SlideGeometry

EMU_PER_INCH = 914400
NS = dict(
    NS,
    c="http://schemas.openxmlformats.org/drawingml/2006/chart",
    r="http://schemas.openxmlformats.org/officeDocument/2006/relationships",
)
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _inferred_rel_path(xml_path: str | Path) -> Optional[str]:
    parts = Path(xml_path).parts
    if "ppt" not in parts:
        return None
    package_parts = parts[parts.index("ppt"):]
    if len(package_parts) != 3 or package_parts[1] != "slides":
        return None
    return f"{package_parts[0]}/slides/_rels/{package_parts[2]}.rels"


def _formatting(text_element: etree._Element) -> dict:
    run = text_element.getparent()
    rpr = run.find("./a:rPr", namespaces=NS) if run is not None else None
    result = {
        "bold": rpr.get("b") in {"1", "true"} if rpr is not None else False,
        "italic": rpr.get("i") in {"1", "true"} if rpr is not None else False,
        "underline": bool(rpr.get("u", "none") not in {"none", ""}) if rpr is not None else False,
        "color": None,
        "font_size": rpr.get("sz") if rpr is not None else None,
    }
    if rpr is not None:
        color = rpr.find("./a:solidFill/a:srgbClr", namespaces=NS)
        if color is not None:
            result["color"] = color.get("val")
    return result


def _runs(elements, tree) -> list[dict]:
    result = []
    for element in elements:
        if element.text is None or not element.text.strip():
            continue
        result.append({
            "text": element.text,
            "formatting": _formatting(element),
            "elem_xpath": tree.getpath(element),
        })
    return result


def _geometry_dict(geometry: SlideGeometry, container) -> dict:
    rect = geometry.rect_for(container)
    if rect is None:
        return {"x": 0.0, "y": 0.0, "w": 0.0, "h": 0.0, "has_geometry": False}
    return {
        "x": rect.x / EMU_PER_INCH,
        "y": rect.y / EMU_PER_INCH,
        "w": rect.width / EMU_PER_INCH,
        "h": rect.height / EMU_PER_INCH,
        "has_geometry": True,
    }


def _related_chart_path(slide_rel_path: str, target: str) -> str:
    # rels находится в _rels, но Target задан относительно исходного slideN.xml.
    slide_path = slide_rel_path.replace("/_rels/", "/")[:-5]
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(slide_path), target))


def _chart_runs(pptx_zip: zipfile.ZipFile, chart_path: str) -> list[dict]:
    root = etree.fromstring(pptx_zip.read(chart_path))
    tree = root.getroottree()
    # Числовые кэши данных не переводятся: они содержат значения диаграммы.
    elements = root.xpath(
        "//a:t | //c:strCache//c:pt/c:v | //c:strLit//c:pt/c:v | //c:ser/c:tx//c:v",
        namespaces=NS,
    )
    unique, seen = [], set()
    for element in elements:
        if id(element) not in seen:
            seen.add(id(element))
            unique.append(element)
    return _runs(unique, tree)


def _chart_data(pptx_zip, container, slide_rel_path, geometry):
    chart_ref = container.find(".//c:chart", namespaces=NS)
    if chart_ref is None or not slide_rel_path:
        return None
    rel_id = chart_ref.get(f"{{{NS['r']}}}id")
    rel_root = etree.fromstring(pptx_zip.read(slide_rel_path))
    relationships = rel_root.xpath(
        "//rel:Relationship[@Id=$rel_id]", namespaces=NS, rel_id=rel_id
    )
    if not relationships:
        return None
    chart_path = _related_chart_path(slide_rel_path, relationships[0].get("Target", ""))
    runs = _chart_runs(pptx_zip, chart_path)
    if not runs:
        return None
    return {
        "text": " | ".join(run["text"] for run in runs),
        "runs": runs,
        "runs_count": len(runs),
        "type": "chart",
        "chart_file": chart_path,
        **_geometry_dict(geometry, container),
    }


def extract_text_blocks(xml_path, pptx_zip=None, slide_rel_path=None):
    """Извлекает текстовые shape, таблицы и текстовые строки диаграмм.

    Каждый run содержит точный XPath, который затем используется для записи.
    """
    if not os.path.exists(xml_path):
        return []
    tree = etree.parse(str(xml_path))
    geometry = SlideGeometry(xml_path)
    slide_rel_path = slide_rel_path or _inferred_rel_path(xml_path)
    blocks = []

    containers = tree.xpath("//p:sp | //p:cxnSp | //p:graphicFrame", namespaces=NS)
    for container in containers:
        try:
            block = None
            if container.tag == f"{{{NS['p']}}}graphicFrame":
                graphic_data = container.find(".//a:graphicData", namespaces=NS)
                uri = graphic_data.get("uri", "") if graphic_data is not None else ""
                table = container.find(".//a:tbl", namespaces=NS)
                if table is not None:
                    elements = table.xpath(".//a:t", namespaces=NS)
                    runs = _runs(elements, tree)
                    if runs:
                        rows = []
                        for tr in table.xpath("./a:tr", namespaces=NS):
                            rows.append([
                                " ".join(tc.xpath(".//a:t/text()", namespaces=NS)).strip()
                                for tc in tr.xpath("./a:tc", namespaces=NS)
                            ])
                        block = {
                            "text": " | ".join(run["text"] for run in runs),
                            "runs": runs,
                            "runs_count": len(runs),
                            "type": "table",
                            "table_data": rows,
                            **_geometry_dict(geometry, container),
                        }
                elif "chart" in uri and pptx_zip is not None:
                    block = _chart_data(pptx_zip, container, slide_rel_path, geometry)
                if block is None:
                    continue
            else:
                elements = container.xpath(".//a:t", namespaces=NS)
                runs = _runs(elements, tree)
                if not runs:
                    continue
                ph = container.find("./p:nvSpPr/p:nvPr/p:ph", namespaces=NS)
                is_title = ph is not None and ph.get("type") in {"title", "ctrTitle"}
                block = {
                    "text": " ".join(run["text"] for run in runs).strip(),
                    "runs": runs,
                    "runs_count": len(runs),
                    "type": "title" if is_title else "shape",
                    "is_title": is_title,
                    **_geometry_dict(geometry, container),
                }
            block["id"] = len(blocks)
            block["xpath"] = tree.getpath(container)
            blocks.append(block)
        except (ValueError, KeyError, etree.XMLSyntaxError, zipfile.BadZipFile) as exc:
            print(f"   ⚠️ Ошибка извлечения блока: {exc}")

    blocks.sort(key=lambda block: (block["y"], block["x"]))
    for index, block in enumerate(blocks):
        block["id"] = index
    return blocks


def split_translation(text: str, runs_count: int):
    """Разбирает [2,1,3] и разделители |, не дублируя текст без разделителей."""
    reorder = None
    match = re.match(r"^\s*\[([\d,]+)\]\s*(.*)$", text, flags=re.DOTALL)
    if match:
        candidate = [int(value) - 1 for value in match.group(1).split(",")]
        if len(candidate) == runs_count and sorted(candidate) == list(range(runs_count)):
            reorder = candidate
        text = match.group(2)

    if "|" not in text:
        parts = [text] if runs_count == 1 else [text] + [""] * (runs_count - 1)
        return parts, reorder

    raw_parts = text.split("|")
    parts = []
    for index, raw in enumerate(raw_parts):
        value = raw.strip()
        if index < len(raw_parts) - 1 and (raw.endswith(" ") or raw_parts[index + 1].startswith(" ")):
            value += " "
        parts.append(value)
    return parts, reorder


def _set_text(element, value):
    element.text = value
    if value.startswith(" ") or value.endswith(" "):
        element.set(XML_SPACE, "preserve")
    else:
        element.attrib.pop(XML_SPACE, None)


def _apply_to_tree(tree, block, new_text):
    runs = block.get("runs", [])
    if not runs:
        return "В блоке нет текстовых элементов"
    parts, reorder = split_translation(new_text, len(runs))
    if len(parts) != len(runs):
        return f"Число частей ({len(parts)}) не совпадает с числом runs ({len(runs)})"

    text_elements = []
    for run in runs:
        xpath = run.get("elem_xpath") or run.get("xpath")
        found = tree.xpath(xpath, namespaces=NS) if xpath else []
        if len(found) != 1:
            return f"XPath не найден или не уникален: {xpath}"
        text_elements.append(found[0])

    source_formatting = []
    if reorder is not None:
        for element in text_elements:
            parent = element.getparent()
            source_formatting.append(copy.deepcopy(parent.find("./a:rPr", namespaces=NS)))

    for index, element in enumerate(text_elements):
        _set_text(element, parts[index])
        if reorder is not None:
            parent = element.getparent()
            old_rpr = parent.find("./a:rPr", namespaces=NS)
            if old_rpr is not None:
                parent.remove(old_rpr)
            source_rpr = source_formatting[reorder[index]]
            if source_rpr is not None:
                parent.insert(0, source_rpr)
    return None


def apply_text_translations(xml_path, translations_by_id, blocks, original_pptx_path, chart_cache):
    """Применяет переводы и возвращает детальный отчёт.

    chart_cache заполняется изменёнными XML диаграмм для последующего экспорта.
    """
    tree = etree.parse(str(xml_path))
    errors, applied = [], 0
    for block in blocks:
        block_id = block["id"]
        new_text = translations_by_id.get(block_id)
        if new_text is None or not str(new_text).strip():
            continue
        target_tree = tree
        if block.get("type") == "chart":
            chart_file = block.get("chart_file")
            if not chart_file:
                errors.append({"block_id": block_id, "error": "Не указан файл диаграммы"})
                continue
            if chart_file not in chart_cache:
                with zipfile.ZipFile(original_pptx_path) as archive:
                    chart_cache[chart_file] = etree.fromstring(archive.read(chart_file))
            target_tree = chart_cache[chart_file].getroottree()
        error = _apply_to_tree(target_tree, block, str(new_text))
        if error:
            errors.append({"block_id": block_id, "error": error})
        else:
            applied += 1

    if applied:
        temporary = str(xml_path) + ".tmp"
        tree.write(temporary, encoding="utf-8", xml_declaration=True, standalone="yes")
        os.replace(temporary, xml_path)
    return {
        "success": applied > 0,
        "applied_blocks": applied,
        "errors": errors,
        "partial": bool(applied and errors),
    }
