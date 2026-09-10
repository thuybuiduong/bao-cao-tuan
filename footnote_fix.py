from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from tempfile import NamedTemporaryFile
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
XML = 'http://www.w3.org/XML/1998/namespace'
NS = {'w': W}


def qn(tag):
    return f'{{{W}}}{tag}'


def ensure_rpr(run):
    rpr = run.find(qn('rPr'))
    if rpr is None:
        rpr = etree.Element(qn('rPr'))
        run.insert(0, rpr)
    return rpr


def force_superscript(run):
    rpr = ensure_rpr(run)
    style = rpr.find(qn('rStyle'))
    if style is None:
        style = etree.SubElement(rpr, qn('rStyle'))
    style.set(qn('val'), 'FootnoteReference')
    va = rpr.find(qn('vertAlign'))
    if va is None:
        va = etree.SubElement(rpr, qn('vertAlign'))
    va.set(qn('val'), 'superscript')


def ensure_footnote_style(styles_root):
    style = styles_root.xpath('./w:style[@w:styleId="FootnoteReference"]', namespaces=NS)
    if style:
        style = style[0]
    else:
        style = etree.Element(qn('style'))
        style.set(qn('type'), 'character')
        style.set(qn('styleId'), 'FootnoteReference')
        name = etree.SubElement(style, qn('name'))
        name.set(qn('val'), 'footnote reference')
        styles_root.append(style)
    rpr = style.find(qn('rPr'))
    if rpr is None:
        rpr = etree.SubElement(style, qn('rPr'))
    va = rpr.find(qn('vertAlign'))
    if va is None:
        va = etree.SubElement(rpr, qn('vertAlign'))
    va.set(qn('val'), 'superscript')


def ensure_bottom_reference(fn):
    fid = fn.get(qn('id'))
    try:
        if int(fid) < 0:
            return
    except Exception:
        return

    p = fn.find(qn('p'))
    if p is None:
        p = etree.SubElement(fn, qn('p'))

    ref_run = None
    for r in p.findall(qn('r')):
        if r.find(qn('footnoteRef')) is not None:
            ref_run = r
            break

    if ref_run is None:
        ref_run = etree.Element(qn('r'))
        etree.SubElement(ref_run, qn('footnoteRef'))
        insert_at = 1 if p.find(qn('pPr')) is not None else 0
        p.insert(insert_at, ref_run)

        spacer = etree.Element(qn('r'))
        t = etree.SubElement(spacer, qn('t'))
        t.set(f'{{{XML}}}space', 'preserve')
        t.text = ' '
        p.insert(insert_at + 1, spacer)

    force_superscript(ref_run)


def fix_footnote_format(docx_path):
    path = Path(docx_path)
    if not path.exists():
        return

    with ZipFile(path, 'r') as zin:
        names = set(zin.namelist())
        docroot = etree.fromstring(zin.read('word/document.xml')) if 'word/document.xml' in names else None
        fnroot = etree.fromstring(zin.read('word/footnotes.xml')) if 'word/footnotes.xml' in names else None
        styles = etree.fromstring(zin.read('word/styles.xml')) if 'word/styles.xml' in names else None

        if docroot is not None:
            for run in docroot.xpath('.//w:r[w:footnoteReference]', namespaces=NS):
                force_superscript(run)

        if fnroot is not None:
            for fn in fnroot.findall(qn('footnote')):
                ensure_bottom_reference(fn)
            for run in fnroot.xpath('.//w:r[w:footnoteRef]', namespaces=NS):
                force_superscript(run)

        if styles is not None:
            ensure_footnote_style(styles)

        replacements = {}
        if docroot is not None:
            replacements['word/document.xml'] = etree.tostring(docroot, xml_declaration=True, encoding='UTF-8', standalone='yes')
        if fnroot is not None:
            replacements['word/footnotes.xml'] = etree.tostring(fnroot, xml_declaration=True, encoding='UTF-8', standalone='yes')
        if styles is not None:
            replacements['word/styles.xml'] = etree.tostring(styles, xml_declaration=True, encoding='UTF-8', standalone='yes')

        with NamedTemporaryFile(delete=False, suffix='.docx') as tf:
            tmp = Path(tf.name)

        try:
            with ZipFile(tmp, 'w', ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = replacements.get(item.filename, zin.read(item.filename))
                    zout.writestr(item, data)
            tmp.replace(path)
        finally:
            if tmp.exists():
                tmp.unlink(missing_ok=True)
