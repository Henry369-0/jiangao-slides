"""List editable PowerPoint objects on each slide, using the standard library."""

import argparse
import posixpath
import sys
from pathlib import PurePosixPath
from urllib.parse import unquote
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


def relationships(archive, part):
    folder, name = posixpath.split(part)
    rels = posixpath.join(folder, "_rels", name + ".rels")
    if rels not in archive.namelist():
        return {}
    result = {}
    for rel in ET.fromstring(archive.read(rels)):
        if rel.get("TargetMode") == "External":
            continue
        target = unquote(rel.attrib["Target"])
        target = posixpath.normpath(posixpath.join(folder, target)).lstrip("/")
        result[rel.attrib["Id"]] = target
    return result


def referenced_id(element, attr="id"):
    # OOXML uses a qualified r:id next to an unqualified numeric id on sldId.
    return next((v for k, v in element.attrib.items() if k.endswith("}" + attr)), None)


def has_embedded_workbook(archive, chart):
    doc = ET.fromstring(archive.read(chart))
    targets = relationships(archive, chart)
    for data in doc.iterfind(".//{*}externalData"):
        path = targets.get(referenced_id(data), "")
        if PurePosixPath(path).suffix.lower() in {".xlsx", ".xlsm", ".xlsb", ".xls"} and path in archive.namelist():
            return True
    return False


def inspect(archive):
    root_rels = relationships(archive, "")
    presentation = next((p for p in root_rels.values() if p.endswith("/presentation.xml")), "ppt/presentation.xml")
    doc = ET.fromstring(archive.read(presentation))
    slide_parts = relationships(archive, presentation)
    for slide_id in doc.iterfind("./{*}sldIdLst/{*}sldId"):
        part = slide_parts[referenced_id(slide_id)]
        doc_slide = ET.fromstring(archive.read(part))
        tree = doc_slide.find("./{*}cSld/{*}spTree")
        if tree is None:
            raise ValueError(f"Missing shape tree in {part}")
        text_shapes = sum(
            any((node.text or "").strip() for node in shape.iterfind("./{*}txBody//{*}t"))
            for shape in tree.iterfind(".//{*}sp")
        )
        pictures = sum(
            referenced_id(blip, "embed") is not None
            for blip in tree.iterfind(".//{*}pic/{*}blipFill/{*}blip")
        )
        slide_rels = relationships(archive, part)
        charts = [slide_rels[referenced_id(c)] for c in tree.iterfind(".//{*}graphicData/{*}chart")]
        tables = sum(1 for _ in tree.iterfind(".//{*}graphicData/{*}tbl"))
        yield text_shapes, pictures, len(charts), sum(has_embedded_workbook(archive, c) for c in charts), tables


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    args = parser.parse_args()
    try:
        with ZipFile(args.pptx) as archive:
            rows = list(inspect(archive))
    except (OSError, BadZipFile, ET.ParseError, ValueError, KeyError) as error:
        print(f"Could not inspect {args.pptx}: {error}", file=sys.stderr)
        return 1
    print("页\t含文字形状\t图片\t原生图表\t含工作簿图表\t原生表格")
    for page, row in enumerate(rows, 1):
        print(page, *row, sep="\t")
    return 0


if __name__ == "__main__":
    sys.exit(main())
