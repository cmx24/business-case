"""PowerPoint-safe packaging steps shared by the deck builds (moved from build_v2.py unchanged)."""


def make_powerpoint_safe(path):
    """Remove the embedded fonts and use Arial.

    The source deck embeds three fonts as GUID-obfuscated TrueType but declares them as application/x-fontdata.
    ECMA-376 requires obfuscated fonts to carry the obfuscatedFont content type; PowerPoint is strict about this and
    LibreOffice is not. DM Sans and Space Grotesk are not installed on most PowerPoint machines, so both become Arial,
    which ships with Office on Windows and Mac (and maps to metric-identical Liberation Sans in the QA renders).
    """
    import re
    import shutil
    import tempfile
    import zipfile

    tmp = tempfile.mktemp(suffix=".pptx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            name = item.filename
            if name.startswith("ppt/fonts/"):
                continue
            data = zin.read(name)
            if name == "ppt/presentation.xml":
                x = data.decode("utf-8")
                x = re.sub(r"<(\w+:)?embeddedFontLst>.*?</(\w+:)?embeddedFontLst>", "", x, flags=re.S)
                x = re.sub(r'\s(embedTrueTypeFonts|saveSubsetFonts)="[^"]*"', "", x)
                x = re.sub(r'(<(\w+:)?sldSz [^>]*?)\stype="[^"]*"', r"\1", x)
                data = x.encode("utf-8")
            elif name == "ppt/_rels/presentation.xml.rels":
                x = data.decode("utf-8")
                x = re.sub(r"<(\w+:)?Relationship [^>]*relationships/font\"[^>]*/>", "", x)
                data = x.encode("utf-8")
            elif name == "[Content_Types].xml":
                x = data.decode("utf-8")
                x = re.sub(r'<(\w+:)?Default Extension="fntdata"[^>]*/>', "", x)
                data = x.encode("utf-8")
            elif name.endswith(".xml") and name.startswith("ppt/"):
                x = data.decode("utf-8")
                x = x.replace('typeface="DM Sans"', 'typeface="Arial"').replace('typeface="Space Grotesk"', 'typeface="Arial"')
                data = x.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
    renumber_presentation_rels(path)


def renumber_presentation_rels(path):
    """Renumber ppt/_rels/presentation.xml.rels so slide relationship IDs ascend in presentation order.

    PowerPoint shows slides in sldIdLst order whatever the IDs are, but tools that list slides by relationship ID (an external
    reviewer did on 2026-10-20) read cloned slides as out of order. Renumbering removes that ambiguity; every r:id in
    presentation.xml is remapped from the same table, so references stay resolved.
    """
    import shutil
    import tempfile
    import zipfile
    from lxml import etree

    RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(path) as z:
        pres = etree.fromstring(z.read("ppt/presentation.xml"))
        rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rel_els = list(rels)
    by_id = {r.get("Id"): r for r in rel_els}
    slide_order = [el.get("{%s}id" % RNS) for el in pres.iter() if etree.QName(el).localname == "sldId"]
    master_ids = [el.get("{%s}id" % RNS) for el in pres.iter() if etree.QName(el).localname in ("sldMasterId", "notesMasterId", "handoutMasterId")]
    ordered = master_ids + slide_order + [r.get("Id") for r in rel_els if r.get("Id") not in master_ids + slide_order]
    mapping = {old: "rId%d" % (i + 1) for i, old in enumerate(ordered)}
    for old, rel in by_id.items():
        rel.set("Id", mapping[old])
    rels[:] = sorted(rels, key=lambda r: int(r.get("Id")[3:]))
    for el in pres.iter():
        for k, v in list(el.attrib.items()):
            if k.startswith("{%s}" % RNS) and v in mapping:
                el.set(k, mapping[v])
    tmp = tempfile.mktemp(suffix=".pptx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "ppt/presentation.xml":
                data = etree.tostring(pres, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif item.filename == "ppt/_rels/presentation.xml.rels":
                data = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
            else:
                data = zin.read(item.filename)
            zout.writestr(item, data)
    shutil.move(tmp, path)
