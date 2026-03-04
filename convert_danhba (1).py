#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DANH BẠ ĐIỆN THOẠI - DOCX to HTML Converter v4
=================================================
STATIC HTML - works on ALL mobile browsers (iOS Safari, Samsung Internet, Chrome)
No JavaScript rendering required - all content pre-rendered by Python.
"""

import sys, os, re, base64, io, unicodedata, html as htmlmod
from docx import Document
from docx.oxml.ns import qn


# ============ UTILS ============

def clean_text(text):
    if not text: return ""
    text = re.sub(r'\n+', ', ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    text = re.sub(r',\s*,', ',', text)
    text = re.sub(r',(?!\s)', ', ', text)
    return text.strip(', ')

def clean_phone(phone):
    if not phone: return ""
    return re.sub(r'\s+', ' ', phone).strip()

def remove_vi(text):
    text = unicodedata.normalize('NFD', text)
    text = re.sub(r'[\u0300-\u036f]', '', text)
    return text.replace('đ', 'd').replace('Đ', 'D').lower()

def is_section_header(name, pos, phone):
    """Detect if a table row is a section header (unit name), not a contact"""
    if phone.strip():
        return False
    if not name.strip():
        return False
    # Remove parenthetical notes like "(05 xã)" for uppercase check
    core = re.sub(r'\([^)]*\)', '', name).strip()
    if not core:
        return False
    letters = [c for c in core if c.isalpha()]
    if not letters:
        return False
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    # >70% uppercase letters and longer than 3 chars = likely a section header
    if upper_ratio >= 0.7 and len(core) > 3:
        return True
    return False

def phone_to_tel(phone):
    if not phone: return ""
    d = re.sub(r'[^\d]', '', phone)
    if d.startswith('0'): d = '+84' + d[1:]
    return 'tel:' + d

def get_initial(name):
    parts = name.strip().split()
    return parts[-1][0].upper() if parts and parts[-1] else '?'

def esc(text):
    return htmlmod.escape(text, quote=True)

def resize_img(data, sz=80):
    try:
        from PIL import Image
        img = Image.open(io.BytesIO(data))
        img.thumbnail((sz, sz), Image.LANCZOS)
        if img.mode in ('RGBA', 'P'): img = img.convert('RGB')
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=65, optimize=True)
        return buf.getvalue()
    except: return data

def img_to_datauri(data):
    if not data: return ""
    b64 = base64.b64encode(data).decode('ascii')
    ct = 'image/jpeg'
    if data[:8] == b'\x89PNG\r\n\x1a\n': ct = 'image/png'
    return "data:%s;base64,%s" % (ct, b64)

def extract_img(cell, doc_part):
    for blip in cell.findall('.//' + qn('a:blip')):
        eid = blip.get(qn('r:embed'))
        if eid:
            try:
                p = doc_part.rels[eid].target_part
                return p.blob
            except: continue
    return None


# ============ DOCX PARSER ============

def parse_docx(filepath, photos_dir=None):
    doc = Document(filepath)
    doc_part = doc.part
    
    # Load external photos
    ext_photos = {}
    if photos_dir and os.path.isdir(photos_dir):
        for fn in os.listdir(photos_dir):
            if fn.lower().endswith(('.jpg','.jpeg','.png','.gif','.webp')):
                nm = os.path.splitext(fn)[0].strip()
                try:
                    with open(os.path.join(photos_dir, fn), 'rb') as f:
                        ext_photos[remove_vi(nm)] = f.read()
                except: pass
    
    categories = []
    cur_cat = None
    cur_subcat = None  # For Heading2/3 intermediate grouping
    cur_dept = None
    photo_count = 0
    
    for el in doc.element.body:
        tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        
        if tag == 'p':
            ft = ''.join(n.text or '' for n in el.iter(qn('w:t'))).strip()
            if not ft: continue
            
            sty = ''
            pPr = el.find(qn('w:pPr'))
            if pPr is not None:
                pS = pPr.find(qn('w:pStyle'))
                if pS is not None: sty = pS.get(qn('w:val'), '')
            
            if sty == 'Heading1':
                cur_cat = {"name": clean_text(ft), "subcats": [], "departments": []}
                categories.append(cur_cat)
                cur_subcat = None
                cur_dept = None
            elif sty == 'Heading2':
                if not cur_cat:
                    cur_cat = {"name": "KHÁC", "subcats": [], "departments": []}
                    categories.append(cur_cat)
                cur_subcat = {"name": clean_text(ft), "departments": []}
                cur_cat["subcats"].append(cur_subcat)
                cur_dept = None
            elif sty in ('Heading3', 'Heading4', 'Heading5'):
                if not cur_cat:
                    cur_cat = {"name": "KHÁC", "subcats": [], "departments": []}
                    categories.append(cur_cat)
                cur_dept = {"name": clean_text(ft), "contacts": []}
                if cur_subcat:
                    cur_subcat["departments"].append(cur_dept)
                else:
                    cur_cat["departments"].append(cur_dept)
        
        elif tag == 'tbl':
            rows = el.findall(qn('w:tr'))
            if not rows: continue
            
            hdr = [''.join(n.text or '' for n in c.iter(qn('w:t'))).strip().upper() 
                   for c in rows[0].findall(qn('w:tc'))]
            hs = ' '.join(hdr)
            if 'TRANG' in hs or 'MỤC LỤC' in hs: continue
            if not any(h in hs for h in ['HỌ VÀ TÊN','CHỨC VỤ','ĐIỆN THOẠI']): continue
            
            cm = {'name':0, 'position':1, 'phone':2, 'note':3}
            for ci, h in enumerate(hdr):
                if 'TÊN' in h or 'HỌ' in h: cm['name'] = ci
                elif 'CHỨC' in h: cm['position'] = ci
                elif 'ĐIỆN' in h or 'SĐT' in h: cm['phone'] = ci
                elif 'GHI' in h or 'CHÚ' in h: cm['note'] = ci
            
            if not cur_dept:
                if not cur_cat:
                    cur_cat = {"name": "KHÁC", "subcats": [], "departments": []}
                    categories.append(cur_cat)
                cur_dept = {"name": "LIÊN HỆ", "contacts": []}
                if cur_subcat:
                    cur_subcat["departments"].append(cur_dept)
                else:
                    cur_cat["departments"].append(cur_dept)
            
            for row in rows[1:]:
                cells = row.findall(qn('w:tc'))
                ct = [''.join(n.text or '' for n in c.iter(qn('w:t'))).strip() for c in cells]
                def gc(i): return ct[i] if i < len(ct) else ""
                
                name = clean_text(gc(cm['name']))
                pos = clean_text(gc(cm['position']))
                phone = clean_phone(gc(cm['phone']))
                note = clean_text(gc(cm.get('note', 99)))
                
                if not name or name.upper() in ('HỌ VÀ TÊN','STT'): continue

                # Detect in-table section headers → create new department
                if is_section_header(name, pos, phone):
                    cur_dept = {"name": clean_text(name), "contacts": []}
                    if cur_subcat:
                        cur_subcat["departments"].append(cur_dept)
                    else:
                        cur_cat["departments"].append(cur_dept)
                    continue
                
                photo_uri = ""
                nn = remove_vi(name)
                if nn in ext_photos:
                    photo_uri = img_to_datauri(resize_img(ext_photos[nn]))
                    photo_count += 1
                else:
                    for c in cells:
                        idata = extract_img(c, doc_part)
                        if idata:
                            photo_uri = img_to_datauri(resize_img(idata))
                            photo_count += 1
                            break
                
                cur_dept["contacts"].append({
                    "name": name, "position": pos,
                    "phone": phone, "note": note, "photo": photo_uri
                })
    
    # Clean empty
    def clean_depts(depts):
        return [d for d in depts if d["contacts"]]
    
    for cat in categories:
        cat["departments"] = clean_depts(cat["departments"])
        for sc in cat["subcats"]:
            sc["departments"] = clean_depts(sc["departments"])
        cat["subcats"] = [s for s in cat["subcats"] if s["departments"]]
    categories = [c for c in categories if c["departments"] or c["subcats"]]
    
    print(f"   {photo_count} ảnh")
    return categories


# ============ HTML GENERATOR (STATIC) ============

def build_search_text(contact, dept_name, cat_name, subcat_name=""):
    """Pre-compute search text for a contact"""
    parts = [contact["name"], contact["position"], dept_name, cat_name]
    if subcat_name: parts.append(subcat_name)
    if contact["note"]: parts.append(contact["note"])
    full = ' '.join(parts)
    return remove_vi(full)

def build_vcard_data(contact, dept_name):
    """Minimal vCard data as HTML-safe attribute"""
    parts = contact["name"].strip().split()
    ln = ' '.join(parts[:-1]) if len(parts) > 1 else ''
    fn = parts[-1] if parts else ''
    phone_intl = ''
    if contact["phone"]:
        d = re.sub(r'[^\d]', '', contact["phone"])
        if d.startswith('0'): d = '+84' + d[1:]
        phone_intl = d
    # Pipe-separated: fn|ln|phone|position|org|note
    return '|'.join([fn, ln, phone_intl, contact["position"], dept_name, contact["note"]])


def render_contact_card(contact, idx, dept_name, cat_name, subcat_name=""):
    search_text = build_search_text(contact, dept_name, cat_name, subcat_name)
    vc_data = esc(build_vcard_data(contact, dept_name))
    has_photo = bool(contact["photo"])
    
    if has_photo:
        avatar = '<div class="av ph" onclick="SP(%d)"><img src="%s" alt=""></div>' % (idx, contact["photo"])
    else:
        avatar = '<div class="av">%s</div>' % esc(get_initial(contact["name"]))
    
    phone_html = ""
    if contact["phone"]:
        phone_html = '''<div class="cc-act">
<a href="%s" class="btn btn-c">&#9742; %s</a>
<button class="btn btn-s" onclick="DL(this)">&#128100;+ Lưu</button>
</div>''' % (phone_to_tel(contact["phone"]), esc(contact["phone"]))
    
    note_html = ""
    if contact["note"]:
        note_html = '<div class="cc-nt">%s</div>' % esc(contact["note"])
    
    hp_attr = ' data-hp="1"' if has_photo else ''
    return '''<div class="cc" data-s="%s" data-v="%s"%s>
<div class="cc-top">%s<div class="cc-info">
<div class="cc-nm">%s</div>
<div class="cc-pos">%s</div>%s
</div></div>%s</div>''' % (
        esc(search_text), vc_data, hp_attr,
        avatar, esc(contact["name"]), esc(contact["position"]),
        note_html, phone_html
    )


def render_department(dept, start_idx, cat_name, subcat_name=""):
    cards = []
    idx = start_idx
    for c in dept["contacts"]:
        cards.append(render_contact_card(c, idx, dept["name"], cat_name, subcat_name))
        idx += 1
    
    count_badge = '<span class="dept-cnt">%d</span>' % len(dept["contacts"])
    
    return '''<details class="dept" data-ds="%s">
<summary class="dept-nm">%s%s</summary>
<div class="c-list">%s</div></details>''' % (
        esc(remove_vi(dept["name"])),
        esc(dept["name"]), count_badge,
        '\n'.join(cards)
    ), idx


def generate_html(categories, title="DANH BẠ ĐIỆN THOẠI", subtitle="LÃNH ĐẠO CÁC CƠ QUAN TỈNH THANH HÓA"):
    total_c = sum(len(d["contacts"]) for c in categories for d in c["departments"]) + \
              sum(len(d["contacts"]) for c in categories for s in c["subcats"] for d in s["departments"])
    total_d = sum(len(c["departments"]) for c in categories) + \
              sum(len(s["departments"]) for c in categories for s in c["subcats"])
    total_p = sum(1 for c in categories for d in c["departments"] for x in d["contacts"] if x["photo"]) + \
              sum(1 for c in categories for s in c["subcats"] for d in s["departments"] for x in d["contacts"] if x["photo"])
    
    # Build all content HTML
    content_parts = []
    filter_tabs = []
    idx = 0
    
    for ci, cat in enumerate(categories):
        cat_html_parts = []
        short_name = cat["name"][:28] + '...' if len(cat["name"]) > 30 else cat["name"]
        filter_tabs.append('<button class="ftab" onclick="FT(this)" data-f="%d">%s</button>' % (ci, esc(short_name)))
        
        # Direct departments
        for dept in cat["departments"]:
            dept_html, idx = render_department(dept, idx, cat["name"])
            cat_html_parts.append(dept_html)
        
        # Sub-categories (Heading2 grouping)
        for sc in cat["subcats"]:
            sc_parts = ['<div class="subcat"><div class="subcat-nm">%s</div>' % esc(sc["name"])]
            for dept in sc["departments"]:
                dept_html, idx = render_department(dept, idx, cat["name"], sc["name"])
                sc_parts.append(dept_html)
            sc_parts.append('</div>')
            cat_html_parts.append('\n'.join(sc_parts))
        
        content_parts.append(
            '<div class="cat" data-ci="%d" data-cs="%s">'
            '<div class="cat-hdr"><div class="cat-ttl">%s</div></div>%s</div>' % (
                ci, esc(remove_vi(cat["name"])),
                esc(cat["name"]), '\n'.join(cat_html_parts)
            )
        )
    
    all_content = '\n'.join(content_parts)
    all_tabs = '\n'.join(filter_tabs)
    
    return '''<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="theme-color" content="#0c1426">
<title>%(title)s</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#0a0f1e;color:#f0f4f8;line-height:1.5;-webkit-text-size-adjust:100%%}
a{color:inherit;text-decoration:none}

.hdr{position:sticky;top:0;z-index:100;padding:12px 16px 0;background:rgba(10,15,30,0.95)}
.hdr-in{max-width:640px;margin:0 auto}
.logo{display:flex;align-items:center;gap:10px;margin-bottom:12px}
.logo-ic{width:40px;height:40px;background:linear-gradient(135deg,#60a5fa,#a78bfa);border-radius:10px;display:flex;align-items:center;justify-content:center;color:#fff;font-size:20px;flex-shrink:0}
.logo h1{font-size:.95rem;font-weight:700;color:#93c5fd}
.logo p{font-size:.68rem;color:#64748b;text-transform:uppercase;letter-spacing:.02em}

.s-wrap{position:relative;margin-bottom:8px}
.s-in{width:100%%;padding:11px 40px 11px 38px;background:#1a2234;border:1.5px solid rgba(255,255,255,0.06);border-radius:999px;color:#f0f4f8;font-family:inherit;font-size:.88rem;outline:none}
.s-in:focus{border-color:#60a5fa;box-shadow:0 0 0 3px rgba(96,165,250,0.15)}
.s-ic{position:absolute;left:4px;top:50%%;transform:translateY(-50%%);color:#64748b;font-size:14px;background:none;border:none;padding:8px 10px;cursor:pointer;-webkit-tap-highlight-color:transparent}
.s-clr{position:absolute;right:8px;top:50%%;transform:translateY(-50%%);width:26px;height:26px;background:#111827;border:none;border-radius:50%%;color:#94a3b8;font-size:14px;display:none;cursor:pointer}

.stats{display:flex;gap:6px;padding-bottom:8px;overflow-x:auto;-webkit-overflow-scrolling:touch}
.stats::-webkit-scrollbar{display:none}
.chip{flex-shrink:0;padding:4px 10px;background:#1a2234;border:1px solid rgba(255,255,255,0.06);border-radius:999px;font-size:.7rem;color:#94a3b8;white-space:nowrap}
.chip b{color:#93c5fd}

.ftabs{display:flex;gap:5px;padding-bottom:10px;overflow-x:auto;-webkit-overflow-scrolling:touch}
.ftabs::-webkit-scrollbar{display:none}
.ftab{flex-shrink:0;padding:6px 12px;background:#1a2234;border:1px solid rgba(255,255,255,0.06);border-radius:999px;font-size:.73rem;font-weight:500;color:#94a3b8;cursor:pointer;white-space:nowrap;-webkit-tap-highlight-color:transparent}
.ftab.on{background:#60a5fa;color:#fff;border-color:#60a5fa}

.main{max-width:640px;margin:0 auto;padding:0 16px 80px}

.cat{margin-bottom:14px}
.cat-hdr{position:sticky;top:105px;z-index:10;padding:6px 0}
.cat-ttl{font-size:.7rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#60a5fa;padding:4px 12px;background:rgba(96,165,250,0.12);border:1px solid rgba(99,179,237,0.2);border-radius:999px;display:inline-block}

.subcat{margin:8px 0;padding-left:0}
.subcat-nm{font-size:.75rem;font-weight:600;color:#fbbf24;padding:6px 4px;border-bottom:1px solid rgba(251,191,36,0.15);margin-bottom:6px}

details.dept{margin-bottom:8px}
details.dept>summary{font-size:.78rem;font-weight:600;color:#94a3b8;padding:6px 4px;display:flex;align-items:center;gap:6px;cursor:pointer;-webkit-tap-highlight-color:transparent;list-style:none}
details.dept>summary::-webkit-details-marker{display:none}
details.dept>summary::before{content:'';width:3px;height:13px;background:linear-gradient(135deg,#34d399,#60a5fa);border-radius:2px;flex-shrink:0}
details.dept>summary::after{content:'\\25B6';margin-left:auto;font-size:.5rem;color:#64748b;transition:transform .2s}
details.dept[open]>summary::after{content:'\\25BC'}
.dept-cnt{font-size:.6rem;color:#64748b;background:rgba(255,255,255,0.05);padding:1px 6px;border-radius:999px}

.c-list{display:flex;flex-direction:column;gap:5px}

.cc{background:#1a2234;border:1px solid rgba(255,255,255,0.06);border-radius:12px;padding:11px;transition:background .15s}
.cc.hide{display:none}

.cc-top{display:flex;align-items:center;gap:10px}
.av{width:44px;height:44px;border-radius:50%%;display:flex;align-items:center;justify-content:center;font-size:.8rem;font-weight:700;color:#fff;flex-shrink:0;overflow:hidden;border:2px solid rgba(96,165,250,0.2)}
.av.ph{background:#1a2234;cursor:pointer}
.av:not(.ph){background:linear-gradient(135deg,#60a5fa,#a78bfa)}
.av img{width:100%%;height:100%%;object-fit:cover;display:block}

.cc-info{flex:1;min-width:0}
.cc-nm{font-size:.88rem;font-weight:600;line-height:1.3}
.cc-pos{font-size:.74rem;color:#94a3b8;line-height:1.4}
.cc-nt{font-size:.68rem;color:#fbbf24;margin-top:2px;font-style:italic}

.cc-act{display:flex;gap:6px;margin-top:9px;padding-top:8px;border-top:1px solid rgba(255,255,255,0.06)}
.btn{flex:1;display:flex;align-items:center;justify-content:center;gap:4px;padding:8px;border-radius:8px;font-family:inherit;font-size:.72rem;font-weight:600;cursor:pointer;border:none;-webkit-tap-highlight-color:transparent}
.btn-c{background:rgba(52,211,153,0.12);color:#34d399;border:1px solid rgba(52,211,153,0.15)}
.btn-s{background:rgba(96,165,250,0.12);color:#93c5fd;border:1px solid rgba(99,179,237,0.15)}
.btn:active{opacity:.7}

.modal{position:fixed;top:0;left:0;right:0;bottom:0;z-index:500;background:rgba(0,0,0,.9);display:none;justify-content:center;align-items:center}
.modal.show{display:flex}
.modal-in{text-align:center;padding:20px}
.modal-in img{width:180px;height:180px;object-fit:cover;border-radius:50%%;border:3px solid rgba(96,165,250,.3);margin-bottom:12px}
.modal-nm{font-size:1.05rem;font-weight:700;margin-bottom:2px}
.modal-ps{font-size:.8rem;color:#94a3b8;margin-bottom:16px}
.modal-x{background:#1a2234;border:1px solid rgba(255,255,255,0.06);color:#94a3b8;padding:7px 22px;border-radius:999px;font-size:.78rem;cursor:pointer}

.sr{font-size:.72rem;color:#fbbf24;padding:4px 8px;background:rgba(251,191,36,0.1);border-radius:999px;display:none}
.sr.show{display:inline-block}

mark{background:rgba(251,191,36,.2);color:#fbbf24;border-radius:2px;padding:0 1px}

.empty{text-align:center;padding:40px 20px;color:#64748b;display:none}
.empty.show{display:block}
.empty p{font-size:.85rem}

.go-top{position:fixed;bottom:20px;right:16px;width:40px;height:40px;background:#60a5fa;border:none;border-radius:50%%;color:#fff;font-size:16px;cursor:pointer;display:none;justify-content:center;align-items:center;box-shadow:0 4px 16px rgba(96,165,250,0.3);z-index:90}
.go-top.show{display:flex}

.toast{position:fixed;bottom:70px;left:50%%;transform:translateX(-50%%);background:#1a2234;border:1px solid rgba(99,179,237,0.2);color:#f0f4f8;padding:8px 16px;border-radius:999px;font-size:.76rem;opacity:0;transition:opacity .3s;z-index:200;white-space:nowrap;pointer-events:none}
.toast.show{opacity:1}

</style>
<script>
var curFilter="all";
var searchTimer=null;

var viMap={};
(function(){
  var from="àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ";
  var to  ="aaaaaaaaaaaaaaaaaeeeeeeeeeeeiiiiiooooooooooooooooouuuuuuuuuuuyyyyyd";
  var fromU=from.toUpperCase();
  var toU=to.toUpperCase();
  for(var i=0;i<from.length;i++){viMap[from[i]]=to[i];viMap[fromU[i]]=toU[i]}
  viMap["đ"]="d";viMap["Đ"]="D";
})();

function rv(s){
  if(!s)return"";
  if(typeof s.normalize==="function")s=s.normalize("NFC");
  s=s.toLowerCase();
  var r="";
  for(var i=0;i<s.length;i++){r+=viMap[s[i]]||s[i]}
  return r;
}

function doSearch(){
  var allCards=document.querySelectorAll(".cc");
  var allDepts=document.querySelectorAll(".dept");
  var allCats=document.querySelectorAll(".cat");
  var si=document.getElementById("si");
  if(!si)return;
  var q=si.value.trim();
  var nq=rv(q);
  var found=0;

  var words=[];
  if(nq){
    var parts=nq.split(/\\s+/);
    for(var w=0;w<parts.length;w++){if(parts[w])words.push(parts[w])}
  }

  for(var i=0;i<allCards.length;i++){
    var card=allCards[i];
    var show=true;

    if(curFilter!=="all"){
      var catEl=card.parentNode;
      while(catEl&&!catEl.getAttribute("data-ci"))catEl=catEl.parentNode;
      if(catEl&&catEl.getAttribute("data-ci")!==curFilter)show=false;
    }

    if(show&&words.length>0){
      var searchText=card.getAttribute("data-s")||"";
      for(var wi=0;wi<words.length;wi++){
        if(searchText.indexOf(words[wi])<0){show=false;break}
      }
    }

    if(show){card.style.display="";found++}
    else{card.style.display="none"}
  }

  for(var d=0;d<allDepts.length;d++){
    var dept=allDepts[d];
    var cards=dept.querySelectorAll(".cc");
    var vis=0;
    for(var c=0;c<cards.length;c++){if(cards[c].style.display!=="none")vis++}
    if(words.length>0){
      dept.style.display=vis>0?"":"none";
      if(vis>0)dept.open=true;
    }else{
      dept.style.display="";
    }
  }

  for(var ci=0;ci<allCats.length;ci++){
    var cat=allCats[ci];
    if(curFilter!=="all"&&cat.getAttribute("data-ci")!==curFilter){
      cat.style.display="none";continue;
    }
    var depts=cat.querySelectorAll(".dept");
    var anyVis=false;
    for(var dd=0;dd<depts.length;dd++){if(depts[dd].style.display!=="none")anyVis=true}
    cat.style.display=anyVis?"":"none";
  }

  var sr=document.getElementById("sr");
  if(sr){
    if(nq){sr.textContent="Tìm thấy: "+found;sr.className="sr show"}
    else{sr.className="sr"}
  }

  var em=document.getElementById("em");
  if(em)em.className=(found===0&&nq)?"empty show":"empty";
}

function SI(){
  var si=document.getElementById("si");
  if(!si)return;
  var sc=document.getElementById("sc");
  if(sc)sc.style.display=si.value.length>0?"block":"none";
  clearTimeout(searchTimer);
  searchTimer=setTimeout(doSearch,200);
}

function SK(e){
  if(e.keyCode===13||e.key==="Enter"){e.preventDefault();doSearch()}
}

function SC(){
  var si=document.getElementById("si");
  if(!si)return;
  si.value="";
  var sc=document.getElementById("sc");
  if(sc)sc.style.display="none";
  doSearch();
  si.focus();
}

function FT(btn){
  var tabs=document.querySelectorAll(".ftab");
  for(var i=0;i<tabs.length;i++)tabs[i].classList.remove("on");
  btn.classList.add("on");
  curFilter=btn.getAttribute("data-f");
  doSearch();
}

function GP(card){
  var img=card.querySelector(".av img");
  return img?img.getAttribute("src"):"";
}

function SP(idx){
  var cards=document.querySelectorAll(".cc");
  if(idx>=cards.length)return;
  var card=cards[idx];
  var ph=GP(card);
  if(!ph)return;
  document.getElementById("mI").src=ph;
  var nm=card.querySelector(".cc-nm");
  var ps=card.querySelector(".cc-pos");
  document.getElementById("mN").textContent=nm?nm.textContent:"";
  document.getElementById("mP").textContent=ps?ps.textContent:"";
  document.getElementById("modal").classList.add("show");
}

function DL(btn){
  var card=btn.parentNode.parentNode;
  var data=card.getAttribute("data-v");
  if(!data)return;
  var parts=data.split("|");
  var fn=parts[0]||"",ln=parts[1]||"",ph=parts[2]||"",ti=parts[3]||"",org=parts[4]||"",nt=parts[5]||"";
  var fullName=(ln?ln+" ":"")+fn;

  var NL="\\r\\n";
  var v="BEGIN:VCARD"+NL+"VERSION:3.0"+NL;
  v+="FN:"+fullName+NL;
  v+="N:"+ln+";"+fn+";;;"+NL;
  if(ph)v+="TEL;TYPE=CELL:"+ph+NL;
  if(ti)v+="TITLE:"+ti+NL;
  if(org)v+="ORG:"+org+NL;
  if(nt)v+="NOTE:"+nt+NL;

  var photoSrc=GP(card);
  if(photoSrc){
    var m=photoSrc.match(/base64,(.+)/);
    if(m){
      var ptype=photoSrc.indexOf("png")>-1?"PNG":"JPEG";
      v+="PHOTO;ENCODING=b;TYPE="+ptype+":"+m[1]+NL;
    }
  }
  v+="END:VCARD";

  var fname=fullName.replace(/\\s+/g,"_")+".vcf";

  if(navigator.share){
    try{
      var file=new File([v],fname,{type:"text/vcard"});
      if(navigator.canShare&&navigator.canShare({files:[file]})){
        navigator.share({files:[file]}).then(function(){
          showToast("Đã chia sẻ: "+fullName);
        }).catch(function(){});
        return;
      }
    }catch(e){}
  }

  try{
    var blob=new Blob([v],{type:"text/vcard"});
    var url=URL.createObjectURL(blob);
    var a=document.createElement("a");
    a.href=url;
    a.download=fname;
    document.body.appendChild(a);
    a.click();
    setTimeout(function(){document.body.removeChild(a);URL.revokeObjectURL(url)},1000);
    showToast("Đã tải: "+fullName);
  }catch(e){
    window.location.href="data:text/vcard;charset=utf-8,"+encodeURIComponent(v);
  }
}

function showToast(msg){
  var t=document.getElementById("toast");
  if(!t)return;
  t.textContent=msg;
  t.classList.add("show");
  clearTimeout(t._t);
  t._t=setTimeout(function(){t.classList.remove("show")},2500);
}

window.onscroll=function(){
  var gt=document.getElementById("gt");
  if(gt)gt.className=window.pageYOffset>400?"go-top show":"go-top";
};
</script>
</head>
<body>

<noscript>
<div style="background:#1e3a5f;color:#93c5fd;padding:10px 16px;text-align:center;font-size:13px;line-height:1.6;border-bottom:1px solid rgba(96,165,250,0.3)">
<b>Huong dan:</b> Bam vao ten don vi de mo/dong danh sach. Dung tim kiem cua trinh xem (&#128269; goc tren) de tim nhanh.
</div>
</noscript>

<div class="hdr"><div class="hdr-in">
<div class="logo">
<div class="logo-ic">&#9742;</div>
<div><h1>%(title)s</h1><p>%(subtitle)s</p></div>
</div>
<div class="s-wrap">
<input type="search" class="s-in" id="si" placeholder="Tìm tên, chức vụ, đơn vị, xã, phường..." autocomplete="off" spellcheck="false" enterkeyhint="search" oninput="SI()" onkeydown="SK(event)">
<button class="s-ic" id="sb" onclick="doSearch()">&#128269;</button>
<button class="s-clr" id="sc" onclick="SC()">&times;</button>
</div>
<div class="stats">
<div class="chip"><b>%(tc)d</b>&nbsp;liên hệ</div>
<div class="chip"><b>%(td)d</b>&nbsp;đơn vị</div>
<div class="chip"><b>%(tp)d</b>&nbsp;có ảnh</div>
<span class="sr" id="sr"></span>
</div>
<div class="ftabs" id="ft">
<button class="ftab on" onclick="FT(this)" data-f="all">Tất cả</button>
%(tabs)s
</div>
</div></div>

<div class="main" id="mc">
%(content)s
<div class="empty" id="em"><p>Không tìm thấy kết quả</p></div>
</div>

<button class="go-top" id="gt" onclick="window.scrollTo(0,0)">&#9650;</button>
<div class="toast" id="toast"></div>

<div class="modal" id="modal" onclick="this.classList.remove('show')">
<div class="modal-in" onclick="event.stopPropagation()">
<img id="mI" src="" alt=""><div class="modal-nm" id="mN"></div><div class="modal-ps" id="mP"></div>
<button class="modal-x" onclick="document.getElementById('modal').classList.remove('show')">Đóng</button>
</div>
</div>

</body>
</html>''' % {
        'title': esc(title),
        'subtitle': esc(subtitle),
        'tc': total_c,
        'td': total_d,
        'tp': total_p,
        'tabs': all_tabs,
        'content': all_content
    }


# ============ MAIN ============

def main():
    if len(sys.argv) < 2:
        print("DANH BẠ v4 - DOCX to HTML (Static, Mobile-safe)")
        print("python %s <input.docx> [output.html] [--photos folder]" % sys.argv[0])
        sys.exit(1)
    
    inp = sys.argv[1]
    if not os.path.exists(inp):
        print("Khong tim thay: " + inp); sys.exit(1)
    
    out, pdir = None, None
    args = sys.argv[2:]
    i = 0
    while i < len(args):
        if args[i] == '--photos' and i+1 < len(args): pdir = args[i+1]; i += 2
        elif not out: out = args[i]; i += 1
        else: i += 1
    if not out: out = os.path.splitext(inp)[0] + '.html'
    
    print("Doc: " + inp)
    cats = parse_docx(inp, pdir)
    tc = sum(len(d["contacts"]) for c in cats for d in c["departments"]) + \
         sum(len(d["contacts"]) for c in cats for s in c["subcats"] for d in s["departments"])
    tp = sum(1 for c in cats for d in c["departments"] for x in d["contacts"] if x["photo"]) + \
         sum(1 for c in cats for s in c["subcats"] for d in s["departments"] for x in d["contacts"] if x["photo"])
    print("%d lien he (%d co anh)" % (tc, tp))
    
    h = generate_html(cats)
    with open(out, 'w', encoding='utf-8') as f: f.write(h)
    sz = os.path.getsize(out)
    u = "KB" if sz < 1048576 else "MB"
    v = sz/1024 if sz < 1048576 else sz/1048576
    print("Xuat: %s (%.0f%s)" % (out, v, u))

    # Generate companion VCF file with all contacts
    vcf_out = os.path.splitext(out)[0] + '.vcf'
    generate_vcf(cats, vcf_out)


def generate_vcf(categories, vcf_path):
    """Generate a single VCF file containing all contacts"""
    NL = "\r\n"
    vcards = []
    count = 0

    def add_contacts(dept, cat_name):
        nonlocal count
        for c in dept["contacts"]:
            parts = c["name"].strip().split()
            ln = ' '.join(parts[:-1]) if len(parts) > 1 else ''
            fn = parts[-1] if parts else ''

            v = "BEGIN:VCARD" + NL + "VERSION:3.0" + NL
            v += "FN:" + c["name"] + NL
            v += "N:" + ln + ";" + fn + ";;;" + NL
            if c["phone"]:
                d = re.sub(r'[^\d]', '', c["phone"])
                if d.startswith('0'): d = '+84' + d[1:]
                v += "TEL;TYPE=CELL:" + d + NL
            if c["position"]:
                v += "TITLE:" + c["position"] + NL
            v += "ORG:" + dept["name"] + NL
            if c["note"]:
                v += "NOTE:" + c["note"] + NL
            # Add photo
            if c["photo"]:
                import re as re2
                m = re2.search(r'base64,(.+)', c["photo"])
                if m:
                    ptype = "PNG" if "png" in c["photo"] else "JPEG"
                    v += "PHOTO;ENCODING=b;TYPE=" + ptype + ":" + m.group(1) + NL
            v += "CATEGORIES:" + cat_name + NL
            v += "END:VCARD"
            vcards.append(v)
            count += 1

    for cat in categories:
        for dept in cat["departments"]:
            add_contacts(dept, cat["name"])
        for sc in cat["subcats"]:
            for dept in sc["departments"]:
                add_contacts(dept, cat["name"] + " - " + sc["name"])

    with open(vcf_path, 'w', encoding='utf-8') as f:
        f.write((NL).join(vcards))

    sz = os.path.getsize(vcf_path)
    u = "KB" if sz < 1048576 else "MB"
    sv = sz/1024 if sz < 1048576 else sz/1048576
    print("VCF: %s (%d lien he, %.0f%s)" % (vcf_path, count, sv, u))


if __name__ == '__main__':
    main()
