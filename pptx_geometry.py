"""Координаты объектов PPTX в системе координат слайда.

Все вычисления выполняются в EMU. Групповые координаты не складываются:
дочерняя система сначала масштабируется через ext/chExt, затем смещается
через off/chOff. Для вложенных групп преобразования компонуются рекурсивно.
"""

from __future__ import annotations

from dataclasses import dataclass
from math import cos, radians, sin
from pathlib import Path
from typing import Optional

from lxml import etree

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
P_GRP_SP = f"{{{NS['p']}}}grpSp"
P_GRAPHIC_FRAME = f"{{{NS['p']}}}graphicFrame"
P_SP = f"{{{NS['p']}}}sp"
A_TBL = f"{{{NS['a']}}}tbl"


@dataclass(frozen=True)
class Matrix:
    """Аффинное преобразование: x'=a*x+c*y+e, y'=b*x+d*y+f."""

    a: float = 1.0
    b: float = 0.0
    c: float = 0.0
    d: float = 1.0
    e: float = 0.0
    f: float = 0.0

    def __matmul__(self, other: "Matrix") -> "Matrix":
        """Возвращает преобразование self(other(point))."""
        return Matrix(
            self.a * other.a + self.c * other.b,
            self.b * other.a + self.d * other.b,
            self.a * other.c + self.c * other.d,
            self.b * other.c + self.d * other.d,
            self.a * other.e + self.c * other.f + self.e,
            self.b * other.e + self.d * other.f + self.f,
        )

    def point(self, x: float, y: float) -> tuple[float, float]:
        return self.a * x + self.c * y + self.e, self.b * x + self.d * y + self.f


@dataclass(frozen=True)
class Rect:
    x: float
    y: float
    width: float
    height: float

    @property
    def corners(self) -> tuple[tuple[float, float], ...]:
        return (
            (self.x, self.y),
            (self.x + self.width, self.y),
            (self.x + self.width, self.y + self.height),
            (self.x, self.y + self.height),
        )


def translation(x: float, y: float) -> Matrix:
    return Matrix(e=x, f=y)


def scaling(x: float, y: float) -> Matrix:
    return Matrix(a=x, d=y)


def rotation_flip(xfrm: etree._Element, center_x: float, center_y: float) -> Matrix:
    """Поворот и отражение вокруг центра объекта/группы."""
    angle = radians(int(xfrm.get("rot", "0")) / 60000.0)
    cos_a, sin_a = cos(angle), sin(angle)
    flip_x = -1.0 if xfrm.get("flipH") in {"1", "true"} else 1.0
    flip_y = -1.0 if xfrm.get("flipV") in {"1", "true"} else 1.0
    transform = Matrix(a=cos_a * flip_x, b=sin_a * flip_x,
                       c=-sin_a * flip_y, d=cos_a * flip_y)
    return translation(center_x, center_y) @ transform @ translation(-center_x, -center_y)


def direct_xfrm(element: etree._Element) -> Optional[etree._Element]:
    """Находит только собственный xfrm, не xfrm вложенных объектов."""
    paths = {
        P_SP: "./p:spPr/a:xfrm",
        f"{{{NS['p']}}}cxnSp": "./p:spPr/a:xfrm",
        f"{{{NS['p']}}}pic": "./p:spPr/a:xfrm",
        P_GRP_SP: "./p:grpSpPr/a:xfrm",
        P_GRAPHIC_FRAME: "./p:xfrm",
    }
    path = paths.get(element.tag)
    return element.find(path, namespaces=NS) if path else None


def local_rect(element: etree._Element) -> Optional[Rect]:
    if element.tag == A_TBL:
        element = next(element.iterancestors(P_GRAPHIC_FRAME), None)
        if element is None:
            return None
    xfrm = direct_xfrm(element)
    if xfrm is None:
        return None
    off = xfrm.find("./a:off", namespaces=NS)
    ext = xfrm.find("./a:ext", namespaces=NS)
    if off is None or ext is None:
        return None
    return Rect(
        float(off.get("x", "0")),
        float(off.get("y", "0")),
        float(ext.get("cx", "0")),
        float(ext.get("cy", "0")),
    )


def group_matrix(group: etree._Element) -> Matrix:
    """Преобразование дочерних координат группы в родительские."""
    xfrm = direct_xfrm(group)
    if xfrm is None:
        return Matrix()
    off = xfrm.find("./a:off", namespaces=NS)
    ext = xfrm.find("./a:ext", namespaces=NS)
    ch_off = xfrm.find("./a:chOff", namespaces=NS)
    ch_ext = xfrm.find("./a:chExt", namespaces=NS)
    if off is None or ext is None:
        return Matrix()

    ox, oy = float(off.get("x", "0")), float(off.get("y", "0"))
    width, height = float(ext.get("cx", "0")), float(ext.get("cy", "0"))
    chx = float(ch_off.get("x", "0")) if ch_off is not None else 0.0
    chy = float(ch_off.get("y", "0")) if ch_off is not None else 0.0
    chw = float(ch_ext.get("cx", "0")) if ch_ext is not None else width
    chh = float(ch_ext.get("cy", "0")) if ch_ext is not None else height
    sx = width / chw if chw else 1.0
    sy = height / chh if chh else 1.0

    base = translation(ox, oy) @ scaling(sx, sy) @ translation(-chx, -chy)
    return rotation_flip(xfrm, ox + width / 2.0, oy + height / 2.0) @ base


def absolute_rect(element: etree._Element, include_own_transform: bool = True) -> Optional[Rect]:
    """Возвращает видимый axis-aligned bounding box в координатах слайда.

    Учитываются все родительские группы, их масштаб, поворот и отражение.
    Собственный поворот/отражение объекта также можно включить в bounding box.
    """
    rect = local_rect(element)
    if rect is None:
        return None

    transform = Matrix()
    if include_own_transform:
        xfrm = direct_xfrm(element)
        if xfrm is not None:
            transform = rotation_flip(
                xfrm, rect.x + rect.width / 2.0, rect.y + rect.height / 2.0
            )

    parent = element.getparent()
    while parent is not None:
        if parent.tag == P_GRP_SP:
            transform = group_matrix(parent) @ transform
        parent = parent.getparent()

    points = [transform.point(x, y) for x, y in rect.corners]
    xs, ys = [p[0] for p in points], [p[1] for p in points]
    return Rect(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))


class SlideGeometry:
    """Геометрия слайда с наследованием placeholder от layout/master."""

    def __init__(self, slide_xml_path: str | Path):
        self.slide_xml_path = Path(slide_xml_path).resolve()
        self._placeholder_cache: dict[int, Optional[Rect]] = {}

    def rect_for(self, element: etree._Element) -> Optional[Rect]:
        if element.tag == A_TBL:
            frame = next(element.iterancestors(P_GRAPHIC_FRAME), None)
            return absolute_rect(frame) if frame is not None else None
        rect = absolute_rect(element)
        if rect is not None or element.tag != P_SP:
            return rect
        key = id(element)
        if key not in self._placeholder_cache:
            self._placeholder_cache[key] = self._inherited_placeholder_rect(element)
        return self._placeholder_cache[key]

    @staticmethod
    def _placeholder(element: etree._Element) -> Optional[etree._Element]:
        return element.find("./p:nvSpPr/p:nvPr/p:ph", namespaces=NS)

    @classmethod
    def _matching_placeholder(cls, root: etree._Element, source: etree._Element):
        ph = cls._placeholder(source)
        if ph is None:
            return None
        idx, ph_type = ph.get("idx"), ph.get("type", "body")
        candidates = root.xpath("//p:sp[p:nvSpPr/p:nvPr/p:ph]", namespaces=NS)
        if idx is not None:
            candidates = [c for c in candidates if cls._placeholder(c).get("idx") == idx]
        else:
            candidates = [
                c for c in candidates
                if cls._placeholder(c).get("idx") is None
                and cls._placeholder(c).get("type", "body") == ph_type
            ]
        return candidates[0] if candidates else None

    @staticmethod
    def _related_part(source: Path, relationship_suffix: str) -> Optional[Path]:
        rels_path = source.parent / "_rels" / f"{source.name}.rels"
        if not rels_path.exists():
            return None
        root = etree.parse(str(rels_path)).getroot()
        for rel in root.xpath("//rel:Relationship", namespaces=NS):
            if rel.get("Type", "").endswith(relationship_suffix):
                target = rel.get("Target")
                if target and not target.startswith("/"):
                    return (source.parent / target).resolve()
        return None

    def _inherited_placeholder_rect(self, element: etree._Element) -> Optional[Rect]:
        if self._placeholder(element) is None:
            return None
        layout_path = self._related_part(self.slide_xml_path, "/slideLayout")
        if layout_path is None or not layout_path.exists():
            return None
        layout_root = etree.parse(str(layout_path)).getroot()
        layout_shape = self._matching_placeholder(layout_root, element)
        if layout_shape is not None:
            rect = absolute_rect(layout_shape)
            if rect is not None:
                return rect
        master_path = self._related_part(layout_path, "/slideMaster")
        if master_path is None or not master_path.exists():
            return None
        master_root = etree.parse(str(master_path)).getroot()
        master_shape = self._matching_placeholder(master_root, layout_shape or element)
        return absolute_rect(master_shape) if master_shape is not None else None
