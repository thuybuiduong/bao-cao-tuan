import re, shutil, tempfile, zipfile
from copy import deepcopy
from pathlib import Path
import openpyxl
from docx import Document
from docx.shared import RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

RED = RGBColor(255,0,0)

TARGETS = {
    'Tin báo': ('hoạt động thực hành quyền công tố, kiểm sát việc tiếp nhận', 16),
    'KSĐT': ('hoạt động thực hành quyền công tố và kiểm sát điều tra', 20),
    'Truy tố': ('hoạt động công tố và kiểm sát hoạt động tư pháp trong giai đoạn truy tố', 17),
    'XXST': ('hoạt động công tố và kiểm sát hoạt động tư pháp trong giai đoạn xét xử', 14),
    'THAHS': ('hoạt động kiểm sát thi hành án hình sự', 39),
    'DS.HNGĐ': ('hoạt động kiểm sát việc giải quyết vụ việc dân sự', 19),
    'KDTM': ('hoạt động kiểm sát việc giải quyết vụ việc kinh doanh', 19),
    'QĐ AD BPXLHC tại Tòa': ('hoạt động kiểm sát giải quyết vụ, việc kinh doanh', 16),
    'HC': ('hoạt động kiểm sát việc giải quyết vụ án hành chính', 18),
    'THADS': ('hoạt động kiểm sát thi hành án dân sự', 17),
    'Khiếu tố': ('* hoạt động kiểm sát:', 27),
}

LABELS = {
    'YC điều tra':'yêu cầu điều tra',
    'YC khởi tố vụ án':'yêu cầu khởi tố vụ án',
    'YC khởi tố bị can':'yêu cầu khởi tố bị can',
    'YC không khởi tố vụ án':'yêu cầu không khởi tố vụ án',
    'YC không khởi tố bị can':'yêu cầu không khởi tố bị can',
    'YC thay đổi quyết định khởi tố vụ án':'yêu cầu thay đổi quyết định khởi tố vụ án',
    'YC thay đổi quyết định khởi tố bị can':'yêu cầu thay đổi quyết định khởi tố bị can',
    'YC phong toả tài khoản':'yêu cầu phong tỏa tài khoản',
    'YC phục hồi án TĐC':'yêu cầu phục hồi vụ án tạm đình chỉ',
    'YC giải quyết vụ án TĐC':'yêu cầu rà soát, giải quyết vụ án tạm đình chỉ',
    'YC truy nã':'yêu cầu truy nã',
    'YC CQĐT chỉ đạo đẩy nhanh các hoạt động điều tra vụ án':'yêu cầu CQĐT đẩy nhanh hoạt động điều tra',
    'YC CQĐT rà soát, giải quyết đối với 01 vụ án tạm đình chỉ đã hết thời hiệu truy cứu trách nhiệm hình sự':'yêu cầu CQĐT rà soát, giải quyết vụ án tạm đình chỉ đã hết thời hiệu truy cứu trách nhiệm hình sự',
    'YC bổ sung QĐ KTVA':'yêu cầu bổ sung quyết định khởi tố vụ án',
    'YC bổ sung QĐ KTBC':'yêu cầu bổ sung quyết định khởi tố bị can',
}

def norm(s):
    return re.sub(r'\s+',' ',str(s or '')).strip().lower()

def is_num(v):
    return isinstance(v,(int,float)) and not isinstance(v,bool)

def fmt(v):
    if v is None: return ''
    if isinstance(v,float) and v.is_integer(): v=int(v)
    return f'{v:,}'.replace(',','.') if isinstance(v,(int,float)) else str(v)

def unit_name(raw):
    s=' '.join(str(raw or '').split())
    if not s or s.lower()=='tổng': return ''
    m=re.match(r'KV\s*(\d+)',s,re.I)
    if m: return f'VKS khu vực {m.group(1)}'
    m=re.match(r'Phòng\s*(\w+)',s,re.I)
    if m: return f'Phòng {m.group(1)}'
    return s.split('\n')[0]

def header_for(ws,c):
    for cc in range(c,0,-1):
        v=ws.cell(1,cc).value
        if v not in (None,''):
            return v
    return ''

def totals_start(ws):
    for c in range(1,ws.max_column+1):
        if norm(ws.cell(1,c).value)=='tổng': return c
    return ws.max_column+1

def activity_rows(ws,start):
    tc=totals_start(ws)
    out=[]
    blank=0
    for r in range(start+1,min(ws.max_row,start+40)+1):
        label=ws.cell(r,1).value
        if label in (None,''):
            blank+=1
            if blank>=2: break
            continue
        blank=0
        label=str(label).strip()
        nlabel=norm(label)
        if nlabel.startswith('vi phạm') or nlabel in {'vi phạm','vi phạm '}:
            continue
        totals=[]
        for c in range(tc,ws.max_column+1):
            v=ws.cell(r,c).value
            if is_num(v): totals.append((ws.cell(2,c).value,v))
        contributors=[]
        for c in range(2,tc):
            v=ws.cell(r,c).value
            if is_num(v) and v!=0:
                u=unit_name(header_for(ws,c))
                if u: contributors.append((u,ws.cell(2,c).value,v))
        # Deduplicate paired-column carryover and keep only meaningful nonzero rows.
        if not totals and not contributors: continue
        if totals:
            measures=[(m,v) for m,v in totals if v not in (None,0)]
        else:
            measures=[]
            bym={}
            for u,m,v in contributors: bym[m]=bym.get(m,0)+v
            measures=list(bym.items())
        if not measures: continue
        out.append((r,label,measures,contributors))
    return out

def footnote_text(contrib, measure=None):
    vals=[]
    seen=set()
    for u,m,v in contrib:
        if measure not in (None,'') and m not in (measure,None,''): continue
        key=(u,v)
        if key in seen: continue
        seen.add(key); vals.append((u,v))
    if not vals: return ''
    # Keep the same concise semicolon-separated style used in the provincial weekly report.
    parts=[]
    for u,v in vals:
        parts.append(f'{u} - {fmt(v)}')
    return '; '.join(parts)+'.'

def metric_text(label,measure,value):
    lab=LABELS.get(label,label)
    n=fmt(value)
    m=norm(measure)
    if norm(label)=='án điểm/ tttg hỏi cung':
        if 'bị can' in m: return ('- Trực tiếp và tham gia hỏi cung ', n, ' bị can.')
        return ('- Xác định án trọng điểm: ', n, ' vụ.')
    unit=''
    if 'bị can' in m: unit=' bị can'
    elif 'bị cáo' in m: unit=' bị cáo'
    elif 'vụ' in m: unit=' vụ'
    elif 'việc' in m: unit=' việc'
    elif 'đơn' in m: unit=' đơn'
    elif 'người' in m: unit=' người'
    elif 'tiền' in m or 'đồng' in m: unit=' đồng'
    return (f'- {lab[:1].upper()+lab[1:]}: ', n, unit+'.')

def find_heading(doc,needle):
    nn=norm(needle).replace('* ','')
    for i,p in enumerate(doc.paragraphs):
        t=norm(p.text).replace('* ','')
        if nn in t: return i
    return None

def block_end(doc,idx):
    for j in range(idx+1,len(doc.paragraphs)):
        t=norm(doc.paragraphs[j].text)
        if not t: continue
        if t.startswith('* vi phạm') or t.startswith('* những biện pháp') or t.startswith('* biện pháp') or re.match(r'^\d+(?:\.\d+)?\.',t):
            return j
    return min(idx+1,len(doc.paragraphs)-1)

def remove_paragraph(p):
    el=p._element; el.getparent().remove(el); p._p=p._element=None

def insert_before(ref,text_parts,marker=None):
    p=ref.insert_paragraph_before()
    p.style=ref.style
    a,n,b=text_parts
    p.add_run(a)
    rr=p.add_run(n); rr.font.color.rgb=RED
    p.add_run(b)
    if marker: p.add_run(marker)
    return p

def add_footnotes(docx_path, mapping):
    if not mapping: return
    src=Path(docx_path)
    tmp=src.with_suffix('.fn.tmp.docx')
    ns='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    W='{'+ns+'}'
    with zipfile.ZipFile(src,'r') as zin:
        document=etree.fromstring(zin.read('word/document.xml'))
        footnotes=etree.fromstring(zin.read('word/footnotes.xml'))
        ids=[]
        for f in footnotes.findall(W+'footnote'):
            try: ids.append(int(f.get(W+'id')))
            except: pass
        next_id=max([x for x in ids if x>=0] or [0])+1
        for marker,text in mapping.items():
            target=None
            for t in document.findall('.//'+W+'t'):
                if marker in (t.text or ''):
                    target=t; break
            if target is None: continue
            run=target.getparent()
            target.text=(target.text or '').replace(marker,'')
            rpr=run.find(W+'rPr')
            if rpr is None:
                rpr=etree.Element(W+'rPr'); run.insert(0,rpr)
            rs=etree.SubElement(rpr,W+'rStyle'); rs.set(W+'val','FootnoteReference')
            va=etree.SubElement(rpr,W+'vertAlign'); va.set(W+'val','superscript')
            fr=etree.SubElement(run,W+'footnoteReference'); fr.set(W+'id',str(next_id))
            fn=etree.SubElement(footnotes,W+'footnote'); fn.set(W+'id',str(next_id))
            p=etree.SubElement(fn,W+'p')
            ppr=etree.SubElement(p,W+'pPr'); ps=etree.SubElement(ppr,W+'pStyle'); ps.set(W+'val','FootnoteText')
            r1=etree.SubElement(p,W+'r'); r1p=etree.SubElement(r1,W+'rPr'); r1s=etree.SubElement(r1p,W+'rStyle'); r1s.set(W+'val','FootnoteReference'); etree.SubElement(r1,W+'footnoteRef')
            r2=etree.SubElement(p,W+'r'); tt=etree.SubElement(r2,W+'t'); tt.set('{http://www.w3.org/XML/1998/namespace}space','preserve'); tt.text=' '+text
            next_id+=1
        with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename=='word/document.xml': data=etree.tostring(document,xml_declaration=True,encoding='UTF-8',standalone='yes')
                elif item.filename=='word/footnotes.xml': data=etree.tostring(footnotes,xml_declaration=True,encoding='UTF-8',standalone='yes')
                else: data=zin.read(item.filename)
                zout.writestr(item,data)
    shutil.move(tmp,src)

def apply_excel_activity_overlay(docx_path,xlsx_path):
    wb=openpyxl.load_workbook(xlsx_path,data_only=True)
    doc=Document(docx_path)
    fnmap={}
    counter=1
    for sname,(needle,start) in TARGETS.items():
        if sname not in wb.sheetnames: continue
        rows=activity_rows(wb[sname],start)
        if not rows: continue
        idx=find_heading(doc,needle)
        if idx is None: continue
        end=block_end(doc,idx)
        ref=doc.paragraphs[end]
        # Remove prior generated/sample activity detail paragraphs, but keep the activity heading itself.
        old=list(doc.paragraphs[idx+1:end])
        for p in old:
            if p._element is not None: remove_paragraph(p)
        for r,label,measures,contrib in rows:
            for measure,value in measures:
                marker=f'[[EXFN{counter}]]'
                ftxt=footnote_text(contrib,measure)
                insert_before(ref,metric_text(label,measure,value),marker if ftxt else None)
                if ftxt: fnmap[marker]=ftxt
                counter+=1
    doc.save(docx_path)
    add_footnotes(docx_path,fnmap)
    return {'activity_footnotes':len(fnmap),'sheets_scanned':len(wb.sheetnames)}
