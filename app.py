#!/usr/bin/env python3
"""
发票收集器 - 单文件版 v2.0
新增功能：
1. 文件打包下载（ZIP）
2. 发票号去重并标注
3. 分类金额汇总
4. 相关文件相邻排列
运行方式: python app.py
然后浏览器打开 http://localhost:8000
"""
import os, re, json, ssl, socket, imaplib, email, hashlib, shutil, tempfile, asyncio, sys, zipfile, io
from pathlib import Path
from datetime import datetime, timedelta
from email.header import decode_header
from copy import copy
from typing import Optional

from fastapi import FastAPI, HTTPException, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse, Response, StreamingResponse
from pydantic import BaseModel

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

app = FastAPI(title="发票收集器", version="2.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ===== Data Models =====
class MailConfig(BaseModel):
    email: str
    auth_code: str
    host: str = "imap.163.com"
    port: int = 993
    date_from: str
    date_to: str

class TaskStatus(BaseModel):
    task_id: str
    status: str
    progress: str
    result: Optional[dict] = None

tasks = {}

# ===== IMAP Helper =====
class StableIMAP4_SSL(imaplib.IMAP4_SSL):
    def _create_socket(self, timeout=None):
        addr_info = socket.getaddrinfo(self.host, self.port, 0, socket.SOCK_STREAM)
        context = self.ssl_context or ssl.create_default_context()
        for family, stype, proto, _, sockaddr in addr_info:
            try:
                sock = socket.socket(family, stype, proto)
                if timeout:
                    sock.settimeout(timeout)
                sock.connect(sockaddr)
                return context.wrap_socket(sock, server_hostname=self.host)
            except:
                continue
        raise ConnectionError('Failed to connect to %s:%d' % (self.host, self.port))

def decode_mime(v):
    if not v:
        return ''
    frags = []
    for part, enc in decode_header(v):
        if isinstance(part, bytes):
            try:
                frags.append(part.decode(enc or 'utf-8', errors='replace'))
            except:
                frags.append(part.decode('utf-8', errors='replace'))
        else:
            frags.append(part)
    return ''.join(frags).strip()

def parse_imap_date(date_str):
    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    dt = datetime.strptime(date_str, '%Y-%m-%d')
    return '%02d-%s-%04d' % (dt.day, months[dt.month-1], dt.year)

# ===== PDF Extraction =====
def parse_pdf(fpath):
    try:
        with pdfplumber.open(str(fpath)) as pdf:
            text = ''
            for p in pdf.pages:
                t = p.extract_text()
                if t:
                    text += t + '\n'
            return text
    except Exception as e:
        return 'ERROR: %s' % e

def extract_seller(text, buyer_name=''):
    m = re.search(r'销\s*名称[：:\s]*([^\n\r]{4,60})', text)
    if m:
        s = re.sub(r'\s+', '', m.group(1).strip())
        s = re.sub(r'(统一社会|纳税人识别号).*$', '', s).strip()
        if len(s) >= 4 and not any(h in s for h in ['规格型号', '单位', '数量', '单价']):
            if not buyer_name or s != buyer_name:
                return s
    m = re.search(r'售\s*名称[：:\s]*([^\n\r]{4,60})', text)
    if m:
        s = re.sub(r'\s+', '', m.group(1).strip())
        s = re.sub(r'(统一社会|纳税人识别号).*$', '', s).strip()
        if len(s) >= 4 and not any(h in s[:8] for h in ['规格型号']):
            if not buyer_name or s != buyer_name:
                return s
    lines = text.split('\n')
    in_sell_section = False
    for i, line in enumerate(lines):
        if '销' in line or '售方' in line:
            in_sell_section = True
        if in_sell_section and re.match(r'^\s*名称[：:\s]', line):
            m2 = re.match(r'\s*名称[：:\s]*(.+)', line)
            if m2:
                s = m2.group(1).strip()
                if len(s) >= 4 and (not buyer_name or s != buyer_name):
                    return s
    all_names = []
    for m3 in re.finditer(r'名称[：:\s]*([^\n\r]{4,50})', text):
        n = re.sub(r'\s+', '', m3.group(1).strip())
        n = re.sub(r'(统一社会|纳税人).*$', '', n).strip()
        if any(h in n[:6] for h in ['规格型号', '单位数量']):
            continue
        if len(n) >= 4 and (not buyer_name or n != buyer_name) and n not in all_names:
            all_names.append(n)
    for n in reversed(all_names):
        if len(n) >= 4:
            return n
    return ''

def extract_date(text):
    m = re.search(r'开票日期[：:\s]*(\d{4})年(\d{2})月(\d{2})日', text)
    if m:
        return '%s-%s-%s' % (m.group(1), m.group(2), m.group(3))
    m2 = re.search(r'开\s*票\s*日\s*期\s*[：:\s]*(\d{4})\s*年\s*(\d{2})\s*月\s*(\d{2})\s*日', text)
    if m2:
        return '%s-%s-%s' % (m2.group(1), m2.group(2), m2.group(3))
    return ''

def extract_amount(text):
    m = re.search(r'价税合计.*?[¥￥]?\s*([\d,]+\.\d{2})', text)
    if m:
        return m.group(1).replace(',', '')
    return ''

def extract_subtotal_tax(text):
    m = re.search(r'合\s*计\s*[¥￥]?\s*([\d,]+\.\d{2})\s*[¥￥]?\s*([\d,]+\.\d{2})', text)
    if m:
        return m.group(1).replace(',', ''), m.group(2).replace(',', '')
    return '', ''

def extract_buyer(text, default_buyer=''):
    m = re.search(r'购\s*名称[：:\s]*([^\n\r销售]{4,40})', text)
    if m:
        b = re.sub(r'\s+', '', m.group(1).strip())
        b = re.sub(r'(统一社会|纳税人).*$', '', b).strip()
        if len(b) >= 4:
            return b
    return default_buyer

def determine_region(seller):
    region_map = {
        '深圳': ['深圳', '南山', '宝安', '福田', '罗湖', '龙华', '龙岗'],
        '广州': ['广州', '宝馔'],
        '河南': ['郑州', '正弘', '洛阳'],
        '海南': ['海南', '华彩'],
        '上海': ['上海', '万程'],
        '湖北': ['湖北', '科壹', '武汉'],
        '天津': ['天津'],
        '北京': ['北京'],
        '成都': ['成都'],
        '杭州': ['杭州'],
        '南京': ['南京'],
        '重庆': ['重庆'],
        '西安': ['西安'],
        '长沙': ['长沙'],
        '苏州': ['苏州'],
        '厦门': ['厦门'],
        '山东': ['山东', '济南', '青岛'],
        '四川': ['四川'],
        '湖南': ['湖南'],
        '广东': ['广东'],
        '福建': ['福建'],
    }
    for region, keywords in region_map.items():
        if any(kw in seller for kw in keywords):
            return region
    return ''

def determine_category(seller, text=''):
    combined = (seller + ' ' + text[:500])
    cat_rules = [
        ('餐饮', ['餐饮', '烧鸟', '灵鹿', '小麦', '欧力给', '博尼塔', '南万喜',
                  '川旺达', '鸟吟', '天纯', '宝馔', '烽源熠鑫', '饭店', '餐厅',
                  '美食', '食府', '火锅', '烧烤', '快餐', '小吃']),
        ('住宿', ['酒店', '威尼斯', '住宿', '宾馆', '旅馆', '民宿', '公寓', '希尔顿',
                  '洲际', '万豪', '香格里拉', '凯悦']),
        ('交通', ['出行', '打车', '滴滴', 'T3', '曹操', 'ETC', '高尔夫', '麻花科技',
                  '汽车租赁', '航空', '铁路', '火车', '高铁', '机票', '航班']),
        ('门票', ['世界之窗', '游览观光', '旅游服务', '景区', '乐园', '门票']),
        ('商超', ['超市', '商场', '正弘', '精华', '便利店', '购物']),
        ('其他', ['玩具', '科壹商贸', '斐乐', '服饰', '体育用品']),
    ]
    for cat, kws in cat_rules:
        if any(kw in combined for kw in kws):
            return cat
    return '其他'

def classify_file(filename, text):
    if '行程单' in filename or ('行程单' in text and '电子发票' not in text):
        return ('supporting', '机票行程单')
    if any(kw in filename for kw in ['课程日程', '酒店安排', '推荐航班', '出行须知', '课程']):
        return ('supporting', '出差行程材料')
    if any(kw in filename for kw in ['入住', '水单', '确认单']):
        return ('supporting', '酒店确认单')
    has_invoice_content = bool(re.search(r'发票号码|价税合计|开票日期', text))
    has_invoice_kw = '发票' in filename or 'dzfp' in filename.lower() or 'qklfp' in filename.lower()
    if has_invoice_content or has_invoice_kw:
        return ('invoice', '发票')
    skip_keywords = ['银行', '对账单', '流水', '签证', 'visa', 'eNoticeLetter',
                     '股票', '基金', '理财', '社保', '公积金', '工资']
    if any(kw in filename for kw in skip_keywords):
        return ('skip', '非报销文件')
    return ('unknown', '未识别')

def extract_uid_from_filename(filename):
    """Extract IMAP UID from saved filename (format: {uid}_{index}_{name})."""
    parts = filename.split('_', 2)
    if len(parts) >= 2 and parts[0].isdigit():
        return parts[0]
    return ''

# ===== Core Collection Logic =====
def collect_invoices(task_id, config):
    work_dir = Path(tempfile.mkdtemp(prefix='invoice_%s_' % task_id))
    raw_dir = work_dir / 'raw'
    raw_dir.mkdir(parents=True, exist_ok=True)
    try:
        tasks[task_id].status = 'running'
        tasks[task_id].progress = '正在连接邮箱...'
        imap = StableIMAP4_SSL(config.host, config.port)
        imap.login(config.email, config.auth_code)
        if '163.com' in config.host:
            imap.xatom('ID', '("name" "invoice-collector" "version" "1.0")')
        tasks[task_id].progress = '邮箱连接成功，正在搜索邮件...'
        imap.select('INBOX', readonly=True)
        since_date = parse_imap_date(config.date_from)
        before_date_dt = datetime.strptime(config.date_to, '%Y-%m-%d') + timedelta(days=1)
        months_list = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        before_date = '%02d-%s-%04d' % (before_date_dt.day, months_list[before_date_dt.month-1], before_date_dt.year)
        search_criteria = '(SINCE %s BEFORE %s)' % (since_date, before_date)
        status, data = imap.uid('search', None, search_criteria)
        if status != 'OK':
            raise Exception('搜索邮件失败: %s' % status)
        uids = data[0].split()
        total_emails = len(uids)
        if total_emails == 0:
            tasks[task_id].status = 'completed'
            tasks[task_id].progress = '未找到邮件'
            tasks[task_id].result = {
                'total_emails': 0, 'invoices': [], 'supporting_docs': [],
                'all_items': [], 'total_amount': 0, 'category_summary': {},
                'no_attach_invoice_emails': []
            }
            imap.logout()
            return
        all_attachments = []
        no_attach_invoice_emails = []
        for i, uid_bytes in enumerate(uids):
            uid = uid_bytes.decode()
            tasks[task_id].progress = '正在处理邮件 (%d/%d)...' % (i+1, total_emails)
            try:
                status, fetch_data = imap.uid('fetch', uid, '(BODY.PEEK[])')
                if status != 'OK':
                    continue
                msg = email.message_from_bytes(fetch_data[0][1])
                subject = decode_mime(msg.get('Subject', ''))
                sender = decode_mime(msg.get('From', ''))
                attachments = []
                html_body = ''
                text_body = ''
                for part in msg.walk():
                    if part.is_multipart():
                        continue
                    ct = part.get_content_type()
                    cd = part.get('Content-Disposition', '')
                    fn = part.get_filename()
                    try:
                        payload = part.get_payload(decode=True)
                        if not payload:
                            continue
                    except:
                        continue
                    if fn or 'attachment' in cd.lower():
                        decoded_fn = decode_mime(fn) if fn else 'attachment_%s' % uid
                        safe_name = re.sub(r'[^\w\u4e00-\u9fff.\-]', '_', decoded_fn)
                        out_name = '%s_%02d_%s' % (uid, len(attachments), safe_name)
                        out_path = raw_dir / out_name
                        counter = 1
                        while out_path.exists():
                            base, ext = out_name.rsplit('.', 1) if '.' in out_name else (out_name, '')
                            out_path = raw_dir / ('%s_%d.%s' % (base, counter, ext)) if ext else raw_dir / ('%s_%d' % (base, counter))
                            counter += 1
                        out_path.write_bytes(payload)
                        attachments.append({
                            'original_filename': decoded_fn,
                            'saved_filename': out_name,
                            'content_type': ct,
                            'size': len(payload),
                        })
                    elif ct == 'text/html':
                        html_body = payload.decode('utf-8', errors='replace')
                    elif ct == 'text/plain':
                        text_body = payload.decode('utf-8', errors='replace')
                urls = re.findall(r'https?://[^\s"\'<>]+', html_body)
                invoice_url_keywords = [
                    'fapiao', 'invoice', 'download', 'pdf', 'fp.', 'dppt', 'bwjf',
                    'chinatax', 'fpjx', 'e-invoice', 'qrcode', 'ewm', 'dzfp'
                ]
                invoice_urls = []
                for u in urls:
                    u_clean = u.rstrip('>"\')')
                    if any(k in u_clean.lower() for k in invoice_url_keywords):
                        invoice_urls.append(u_clean)
                if attachments:
                    all_attachments.extend([dict(a, uid=uid, subject=subject) for a in attachments])
                else:
                    combined = (subject + text_body + html_body).lower()
                    has_invoice_kw = any(kw in combined for kw in
                        ['发票', '电子票', '行程', '报销', '凭证', '酒店', '机票', 'invoice',
                         '行程单', '水单', '入住', '打车', '出行', '美团', '饿了么',
                         '携程', '飞猪', '去哪儿', '同程', '航空', '高铁', '火车'])
                    if has_invoice_kw or invoice_urls:
                        no_attach_invoice_emails.append({
                            'uid': uid, 'subject': subject, 'from': sender,
                            'invoice_urls': invoice_urls[:10],
                        })
            except:
                continue
        imap.logout()
        tasks[task_id].progress = '已下载 %d 个附件，正在解析发票...' % len(all_attachments)
        invoices = []
        supporting_docs = []
        seen_inv_nos = {}  # inv_no -> first invoice dict
        for f in sorted(raw_dir.glob('*')):
            fname = f.name
            fpath = str(f)
            file_uid = extract_uid_from_filename(fname)
            if f.suffix.lower() == '.pdf':
                text = parse_pdf(fpath)
                if text.startswith('ERROR'):
                    continue
                file_type, subtype = classify_file(fname, text)
                if file_type == 'skip':
                    continue
                elif file_type == 'supporting':
                    supporting_docs.append({
                        'source_file': fname, 'source_path': fpath,
                        'type': subtype, 'row_type': '辅助单据',
                        'file_uid': file_uid, 'text_snippet': text[:800],
                    })
                    continue
                elif file_type == 'unknown':
                    continue
                inv_no_m = re.search(r'发票号码[：:\s]*(\d+)', text)
                if not inv_no_m:
                    inv_no_m = re.search(r'(\d{20})', text)
                inv_no = inv_no_m.group(1) if inv_no_m else ''
                is_duplicate = False
                duplicate_of_inv = None
                if inv_no and inv_no in seen_inv_nos:
                    is_duplicate = True
                    duplicate_of_inv = seen_inv_nos[inv_no]
                buyer = extract_buyer(text)
                seller = extract_seller(text, buyer)
                date_val = extract_date(text)
                amount = extract_amount(text)
                no_tax, tax = extract_subtotal_tax(text)
                if not seller:
                    fm = re.search(r'_([^_]+(?:有限公司|个体工商户|店|院))', fname)
                    if fm:
                        seller = fm.group(1)
                if not amount:
                    fa = re.search(r'(\d+\.\d{2})元?', fname)
                    if fa:
                        amount = fa.group(1)
                if not date_val:
                    fd = re.search(r'(\d{4}).(\d{2}).(\d{2})', fname)
                    if fd:
                        date_val = '%s-%s-%s' % (fd.group(1), fd.group(2), fd.group(3))
                category = determine_category(seller, text)
                region = determine_region(seller)
                inv = {
                    'invoice_no': inv_no, 'date': date_val, 'amount': amount,
                    'amount_no_tax': no_tax, 'tax': tax, 'buyer': buyer,
                    'seller': seller, 'region': region, 'category': category,
                    'source_file': fname, 'source_path': fpath,
                    'row_type': '发票', 'file_uid': file_uid,
                    'is_duplicate': is_duplicate, 'duplicate_of': duplicate_of_inv,
                    'text_snippet': text[:800],
                }
                invoices.append(inv)
                if inv_no and not is_duplicate:
                    seen_inv_nos[inv_no] = inv
            elif f.suffix.lower() in ['.png', '.jpg', '.jpeg']:
                if any(kw in fname for kw in ['发票', '二维码', '截图']):
                    supporting_docs.append({
                        'source_file': fname, 'source_path': fpath,
                        'type': '发票截图', 'row_type': '待确认',
                        'file_uid': file_uid, 'text_snippet': '',
                    })
        # Sort invoices by date+amount
        invoices.sort(key=lambda x: (x.get('date', '9999-99-99'),
                                     float(x.get('amount', '0') or 0)))
        # === Interleave: place supporting docs next to matching invoices ===
        ordered_items = []
        used_supporting = set()
        for inv in invoices:
            ordered_items.append(('invoice', inv))
            inv_uid = inv.get('file_uid', '')
            inv_seller = inv.get('seller', '')
            inv_cat = inv.get('category', '')
            for j, sd in enumerate(supporting_docs):
                if j in used_supporting:
                    continue
                sd_uid = sd.get('file_uid', '')
                sd_type = sd.get('type', '')
                matched = False
                # Rule 1: same email UID
                if sd_uid and inv_uid and sd_uid == inv_uid:
                    matched = True
                # Rule 2: hotel confirmation matches hotel invoice by seller name
                elif sd_type == '酒店确认单' and inv_cat == '住宿':
                    sd_text = sd.get('text_snippet', '')
                    if inv_seller:
                        seller_keywords = [inv_seller[i:i+3] for i in range(0, min(len(inv_seller), 8), 2)]
                        if any(kw in sd_text for kw in seller_keywords if len(kw) >= 2):
                            matched = True
                        else:
                            hotel_invs = [i for i in invoices if i.get('category') == '住宿']
                            hotel_sds = [(k, s) for k, s in enumerate(supporting_docs)
                                         if s.get('type') == '酒店确认单' and k not in used_supporting]
                            if len(hotel_invs) == 1 and len(hotel_sds) == 1:
                                matched = True
                # Rule 3: itinerary matches transportation invoice
                elif sd_type == '机票行程单' and inv_cat == '交通':
                    matched = True
                # Rule 4: travel materials match any invoice from same date range
                elif sd_type == '出差行程材料':
                    if inv_uid and sd_uid and inv_uid == sd_uid:
                        matched = True
                if matched:
                    ordered_items.append(('supporting', sd))
                    used_supporting.add(j)
        # Add remaining unmatched supporting docs at the end
        for j, sd in enumerate(supporting_docs):
            if j not in used_supporting:
                ordered_items.append(('supporting', sd))
        # Number all items sequentially
        for i, (item_type, item) in enumerate(ordered_items):
            item['seq_num'] = '%03d' % (i + 1)
        # Update duplicate references with seq numbers
        for inv in invoices:
            if inv.get('is_duplicate') and inv.get('duplicate_of'):
                inv['duplicate_of_seq'] = inv['duplicate_of'].get('seq_num', '')
            else:
                inv['duplicate_of_seq'] = ''
        # === Compute category summary (excluding duplicates) ===
        category_summary = {}
        for inv in invoices:
            if inv.get('is_duplicate'):
                continue
            cat = inv.get('category', '其他')
            try:
                amt = float(inv.get('amount', 0) or 0)
            except:
                amt = 0
            category_summary[cat] = round(category_summary.get(cat, 0) + amt, 2)
        duplicate_count = sum(1 for inv in invoices if inv.get('is_duplicate'))
        # === Generate Excel ===
        tasks[task_id].progress = '正在生成Excel表格...'
        wb = Workbook()
        ws = wb.active
        ws.title = '发票信息'
        headers = ['序号','文件类型','地区','类目','发票日期','发票号码',
                   '不含税金额','税额','金额','购买方名称','销售方名称',
                   '对应发票序号','原文件名','新文件名','关联方式','备注']
        header_font = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        for col, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        data_font = Font(name='微软雅黑', size=10)
        data_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        center_alignment = Alignment(horizontal='center', vertical='center')
        dup_fill = PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid')
        supp_fill = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')
        inv_fill = PatternFill(start_color='F0FDF4', end_color='F0FDF4', fill_type='solid')
        all_rows = []
        for item_type, item in ordered_items:
            if item_type == 'invoice':
                inv = item
                ext = Path(inv['source_file']).suffix.lower()
                cat = inv.get('category', '')
                dt = inv.get('date', '')
                amt = inv.get('amount', '')
                nfn = '%s-%s-%s-%s%s' % (inv['seq_num'], cat, dt, amt, ext) if dt and amt else '%s-%s%s' % (inv['seq_num'], cat, ext)
                amt_val = amt
                try:
                    amt_val = float(amt)
                except:
                    pass
                remark = ''
                if inv.get('is_duplicate'):
                    remark = '⚠️重复发票，与序号%s重复' % inv.get('duplicate_of_seq', '')
                row_data = [
                    inv['seq_num'], '发票', inv['region'], cat, dt,
                    inv['invoice_no'], inv.get('amount_no_tax', ''),
                    inv.get('tax', ''), amt_val, inv['buyer'], inv['seller'],
                    inv['seq_num'], inv['source_file'], nfn, '', remark
                ]
                all_rows.append({'nfn': nfn, 'data': row_data, 'type': 'invoice',
                                 'source': inv, 'is_duplicate': inv.get('is_duplicate', False)})
            else:
                sd = item
                ext = Path(sd['source_file']).suffix.lower()
                cat = sd.get('type', '')
                nfn = '%s-%s%s' % (sd['seq_num'], cat, ext)
                related_seq = ''
                sd_uid = sd.get('file_uid', '')
                for item_type2, item2 in ordered_items:
                    if item_type2 == 'invoice':
                        i_uid = item2.get('file_uid', '')
                        if sd_uid and i_uid and sd_uid == i_uid:
                            related_seq = item2['seq_num']
                            break
                remark = ''
                relation = ''
                if related_seq:
                    remark = '关联发票序号%s' % related_seq
                    relation = '关联'
                row_data = [
                    sd['seq_num'], sd.get('row_type', '辅助单据'), '', '', '',
                    '', '', '', '', '', '',
                    related_seq, sd['source_file'], nfn, relation, remark
                ]
                all_rows.append({'nfn': nfn, 'data': row_data, 'type': 'supporting', 'source': sd})
        for i, r in enumerate(all_rows):
            row = i + 2
            for col, val in enumerate(r['data'], 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.font = data_font
                cell.alignment = center_alignment if col in [1,3,4,5,12,15] else data_alignment
                cell.border = thin_border
                if r.get('is_duplicate'):
                    cell.fill = dup_fill
                elif r['type'] == 'invoice':
                    cell.fill = inv_fill
                elif r['type'] == 'supporting':
                    cell.fill = supp_fill
        col_widths = [6, 10, 8, 8, 12, 24, 12, 10, 12, 22, 30, 10, 30, 30, 8, 24]
        for i, w in enumerate(col_widths):
            ws.column_dimensions[chr(65 + i)].width = w
        # === Add category summary sheet ===
        ws2 = wb.create_sheet('分类汇总')
        ws2.cell(row=1, column=1, value='类别').font = header_font
        ws2.cell(row=1, column=2, value='金额(¥)').font = header_font
        ws2.cell(row=1, column=3, value='占比').font = header_font
        ws2.cell(row=1, column=4, value='发票数').font = header_font
        for c in range(1, 5):
            ws2.cell(row=1, column=c).fill = header_fill
            ws2.cell(row=1, column=c).alignment = header_alignment
            ws2.cell(row=1, column=c).border = thin_border
        total_cat = sum(category_summary.values())
        inv_count_by_cat = {}
        for inv in invoices:
            if inv.get('is_duplicate'):
                continue
            cat = inv.get('category', '其他')
            inv_count_by_cat[cat] = inv_count_by_cat.get(cat, 0) + 1
        for i, (cat, amt) in enumerate(sorted(category_summary.items(), key=lambda x: -x[1])):
            r = i + 2
            ws2.cell(row=r, column=1, value=cat).border = thin_border
            ws2.cell(row=r, column=2, value=amt).border = thin_border
            ws2.cell(row=r, column=2).number_format = '#,##0.00'
            pct = '%.1f%%' % (amt / total_cat * 100) if total_cat > 0 else '0%'
            ws2.cell(row=r, column=3, value=pct).border = thin_border
            ws2.cell(row=r, column=4, value=inv_count_by_cat.get(cat, 0)).border = thin_border
        total_row = len(category_summary) + 2
        ws2.cell(row=total_row, column=1, value='合计').font = Font(name='微软雅黑', bold=True, size=11)
        ws2.cell(row=total_row, column=2, value=total_cat).font = Font(name='微软雅黑', bold=True, size=11)
        ws2.cell(row=total_row, column=2).number_format = '#,##0.00'
        ws2.cell(row=total_row, column=3, value='100%').font = Font(name='微软雅黑', bold=True, size=11)
        ws2.cell(row=total_row, column=4, value=sum(inv_count_by_cat.values())).font = Font(name='微软雅黑', bold=True, size=11)
        for c in range(1, 5):
            ws2.cell(row=total_row, column=c).border = thin_border
        ws2.column_dimensions['A'].width = 16
        ws2.column_dimensions['B'].width = 16
        ws2.column_dimensions['C'].width = 10
        ws2.column_dimensions['D'].width = 10
        xlsx_path = work_dir / '发票信息统计.xlsx'
        wb.save(str(xlsx_path))
        # Copy files to processed dir with new names
        processed_dir = work_dir / 'processed'
        processed_dir.mkdir(parents=True, exist_ok=True)
        for r in all_rows:
            src_name = r['data'][12]
            new_name = r['nfn']
            src_path = raw_dir / src_name if src_name else None
            if src_path and src_path.exists():
                dst = processed_dir / new_name
                c = 1
                while dst.exists():
                    base, ext = new_name.rsplit('.', 1) if '.' in new_name else (new_name, '')
                    dst = processed_dir / ('%s_%d.%s' % (base, c, ext)) if ext else processed_dir / ('%s_%d' % (base, c))
                    c += 1
                shutil.copy2(str(src_path), str(dst))
        # Build result
        total_amount = sum(float(inv.get('amount', 0) or 0) for inv in invoices if not inv.get('is_duplicate'))
        # Build all_items for frontend display (interleaved)
        all_items = []
        for item_type, item in ordered_items:
            if item_type == 'invoice':
                inv = item
                remark = ''
                if inv.get('is_duplicate'):
                    remark = '⚠️重复发票，与序号%s重复' % inv.get('duplicate_of_seq', '')
                all_items.append({
                    'seq': inv['seq_num'],
                    'item_type': 'invoice',
                    'is_duplicate': inv.get('is_duplicate', False),
                    'duplicate_of_seq': inv.get('duplicate_of_seq', ''),
                    'region': inv['region'],
                    'category': inv['category'],
                    'date': inv['date'],
                    'invoice_no': inv['invoice_no'],
                    'amount': inv['amount'],
                    'amount_no_tax': inv.get('amount_no_tax', ''),
                    'tax': inv.get('tax', ''),
                    'buyer': inv['buyer'],
                    'seller': inv['seller'],
                    'remark': remark,
                    'source_file': inv['source_file'],
                })
            else:
                sd = item
                related_seq = ''
                sd_uid = sd.get('file_uid', '')
                for item_type2, item2 in ordered_items:
                    if item_type2 == 'invoice':
                        i_uid = item2.get('file_uid', '')
                        if sd_uid and i_uid and sd_uid == i_uid:
                            related_seq = item2['seq_num']
                            break
                remark = ''
                if related_seq:
                    remark = '关联发票 #%s' % related_seq
                all_items.append({
                    'seq': sd['seq_num'],
                    'item_type': 'supporting',
                    'is_duplicate': False,
                    'duplicate_of_seq': '',
                    'region': '',
                    'category': sd.get('type', ''),
                    'date': '',
                    'invoice_no': '',
                    'amount': '',
                    'amount_no_tax': '',
                    'tax': '',
                    'buyer': '',
                    'seller': '',
                    'remark': remark,
                    'source_file': sd['source_file'],
                })
        result = {
            'task_id': task_id,
            'total_emails': total_emails,
            'total_attachments': len(all_attachments),
            'invoice_count': len(invoices),
            'duplicate_count': duplicate_count,
            'supporting_count': len(supporting_docs),
            'total_amount': round(total_amount, 2),
            'category_summary': category_summary,
            'all_items': all_items,
            'invoices': [
                {
                    'seq': inv['seq_num'], 'invoice_no': inv['invoice_no'],
                    'date': inv['date'], 'amount': inv['amount'],
                    'amount_no_tax': inv.get('amount_no_tax', ''),
                    'tax': inv.get('tax', ''),
                    'buyer': inv['buyer'], 'seller': inv['seller'],
                    'region': inv['region'], 'category': inv['category'],
                    'is_duplicate': inv.get('is_duplicate', False),
                    'duplicate_of_seq': inv.get('duplicate_of_seq', ''),
                    'remark': '⚠️重复发票，与序号%s重复' % inv.get('duplicate_of_seq', '') if inv.get('is_duplicate') else '',
                } for inv in invoices
            ],
            'supporting_docs': [
                {
                    'seq': sd['seq_num'], 'type': sd.get('type', ''),
                    'source_file': sd['source_file'],
                } for sd in supporting_docs
            ],
            'no_attach_invoice_emails': [
                {
                    'uid': e['uid'], 'subject': e['subject'],
                    'from': e['from'], 'invoice_urls': e.get('invoice_urls', []),
                } for e in no_attach_invoice_emails[:20]
            ],
            'work_dir': str(work_dir),
        }
        tasks[task_id].status = 'completed'
        tasks[task_id].progress = '完成！共 %d 张发票（含 %d 张重复），%d 个辅助文件，合计 ¥%.2f' % (
            len(invoices), duplicate_count, len(supporting_docs), total_amount)
        tasks[task_id].result = result
    except Exception as e:
        tasks[task_id].status = 'failed'
        tasks[task_id].progress = '错误: %s' % str(e)
        import traceback
        traceback.print_exc()

# ===== API Endpoints =====
@app.get("/api/health")
async def health():
    return {"status": "ok", "version": "2.0.0"}

@app.post("/api/collect")
async def start_collection(config: MailConfig, background_tasks: BackgroundTasks):
    task_id = hashlib.md5(
        ("%s%s%s%s" % (config.email, config.date_from, config.date_to, datetime.now().isoformat())).encode()
    ).hexdigest()[:12]
    tasks[task_id] = TaskStatus(
        task_id=task_id, status='pending',
        progress='任务已创建，等待执行...', result=None
    )
    background_tasks.add_task(collect_invoices, task_id, config)
    return {"task_id": task_id, "message": "任务已创建"}

@app.get("/api/status/{task_id}")
async def get_status(task_id: str):
    if task_id not in tasks:
        raise HTTPException(status_code=404, detail="任务不存在")
    return tasks[task_id]

@app.get("/api/download/{task_id}")
async def download_excel(task_id: str):
    if task_id not in tasks:
        raise HTTPException(status_code=404, detail="任务不存在")
    task = tasks[task_id]
    if task.status != 'completed' or not task.result:
        raise HTTPException(status_code=400, detail="任务未完成或无结果")
    work_dir = Path(task.result['work_dir'])
    xlsx_path = work_dir / '发票信息统计.xlsx'
    if not xlsx_path.exists():
        raise HTTPException(status_code=404, detail="Excel文件不存在")
    return FileResponse(
        str(xlsx_path),
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename='发票信息统计_%s.xlsx' % task_id
    )

@app.get("/api/download-all/{task_id}")
async def download_all_files(task_id: str):
    """Download all files as a ZIP archive, including the Excel summary."""
    if task_id not in tasks:
        raise HTTPException(status_code=404, detail="任务不存在")
    task = tasks[task_id]
    if task.status != 'completed' or not task.result:
        raise HTTPException(status_code=400, detail="任务未完成或无结果")
    work_dir = Path(task.result['work_dir'])
    processed_dir = work_dir / 'processed'
    if not processed_dir.exists():
        raise HTTPException(status_code=404, detail="文件目录不存在")
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        xlsx_path = work_dir / '发票信息统计.xlsx'
        if xlsx_path.exists():
            zf.write(str(xlsx_path), '发票信息统计.xlsx')
        for f in sorted(processed_dir.glob('*')):
            zf.write(str(f), f.name)
    zip_buffer.seek(0)
    return StreamingResponse(
        zip_buffer,
        media_type='application/zip',
        headers={'Content-Disposition': 'attachment; filename*=UTF-8\'\'%s' % ('发票文件包_%s.zip' % task_id)}
    )

@app.get("/api/files/{task_id}")
async def list_files(task_id: str):
    if task_id not in tasks:
        raise HTTPException(status_code=404, detail="任务不存在")
    task = tasks[task_id]
    if task.status != 'completed' or not task.result:
        raise HTTPException(status_code=400, detail="任务未完成")
    work_dir = Path(task.result['work_dir'])
    processed_dir = work_dir / 'processed'
    if not processed_dir.exists():
        return {"files": []}
    files = []
    for f in sorted(processed_dir.glob('*')):
        files.append({'name': f.name, 'size': f.stat().st_size})
    return {"files": files}

@app.get("/api/file/{task_id}/{filename:path}")
async def download_file(task_id: str, filename: str):
    if task_id not in tasks:
        raise HTTPException(status_code=404, detail="任务不存在")
    task = tasks[task_id]
    if task.status != 'completed' or not task.result:
        raise HTTPException(status_code=400, detail="任务未完成")
    work_dir = Path(task.result['work_dir'])
    file_path = work_dir / 'processed' / filename
    if not file_path.exists():
        file_path = work_dir / 'raw' / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(str(file_path), filename=filename)

# ===== Embedded Frontend =====
HTML_CONTENT = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>发票收集器 - 智能发票管理平台</title>
<style>
:root{--primary:#4F46E5;--primary-hover:#4338CA;--primary-light:#EEF2FF;--success:#059669;--success-light:#D1FAE5;--warning:#D97706;--warning-light:#FEF3C7;--danger:#DC2626;--danger-light:#FEE2E2;--gray-50:#F9FAFB;--gray-100:#F3F4F6;--gray-200:#E5E7EB;--gray-300:#D1D5DB;--gray-400:#9CA3AF;--gray-500:#6B7280;--gray-600:#4B5563;--gray-700:#374151;--gray-800:#1F2937;--gray-900:#111827;--shadow-lg:0 10px 15px -3px rgba(0,0,0,0.1),0 4px 6px -4px rgba(0,0,0,0.1);--radius:12px;--radius-sm:8px}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang SC','Hiragino Sans GB','Microsoft YaHei',sans-serif;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);min-height:100vh;color:var(--gray-800)}
.container{max-width:1100px;margin:0 auto;padding:24px 16px}
.header{text-align:center;padding:40px 0 32px;color:#fff}
.header h1{font-size:36px;font-weight:700;letter-spacing:-.5px;margin-bottom:8px}
.header p{font-size:16px;opacity:.85}
.card{background:#fff;border-radius:var(--radius);box-shadow:var(--shadow-lg);padding:32px;margin-bottom:24px;transition:transform .2s}
.card:hover{transform:translateY(-1px)}
.card-title{font-size:20px;font-weight:600;color:var(--gray-800);margin-bottom:24px;display:flex;align-items:center;gap:10px}
.card-title .icon{width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px}
.icon-blue{background:var(--primary-light);color:var(--primary)}
.icon-green{background:var(--success-light);color:var(--success)}
.icon-orange{background:var(--warning-light);color:var(--warning)}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}
@media(max-width:640px){.form-grid{grid-template-columns:1fr}}
.form-group{display:flex;flex-direction:column;gap:6px}
.form-group.full-width{grid-column:1/-1}
.form-label{font-size:14px;font-weight:500;color:var(--gray-600)}
.form-input,.form-select{padding:12px 16px;border:1.5px solid var(--gray-200);border-radius:var(--radius-sm);font-size:15px;color:var(--gray-800);transition:border-color .2s,box-shadow .2s;outline:0;background:#fff}
.form-input:focus,.form-select:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(79,70,229,.1)}
.form-input::placeholder{color:var(--gray-400)}
.form-hint{font-size:12px;color:var(--gray-400)}
.btn{padding:14px 32px;border:0;border-radius:var(--radius-sm);font-size:16px;font-weight:600;cursor:pointer;transition:all .2s;display:inline-flex;align-items:center;gap:8px}
.btn-primary{background:var(--primary);color:#fff}
.btn-primary:hover{background:var(--primary-hover);transform:translateY(-1px)}
.btn-primary:disabled{background:var(--gray-300);cursor:not-allowed;transform:none}
.btn-success{background:var(--success);color:#fff}
.btn-success:hover{opacity:.9;transform:translateY(-1px)}
.btn-outline{background:#fff;color:var(--primary);border:1.5px solid var(--primary)}
.btn-outline:hover{background:var(--primary-light)}
.progress-bar-container{background:var(--gray-100);border-radius:99px;height:10px;overflow:hidden;margin-bottom:12px}
.progress-bar{height:100%;background:linear-gradient(90deg,var(--primary),#818CF8);border-radius:99px;transition:width .5s ease;width:0}
.progress-text{font-size:14px;color:var(--gray-600);text-align:center}
.spinner{display:inline-block;width:36px;height:36px;border:3px solid rgba(79,70,229,.2);border-top-color:var(--primary);border-radius:50%;animation:spin .8s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.stats-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:16px;margin-bottom:24px}
@media(max-width:640px){.stats-grid{grid-template-columns:repeat(2,1fr)}}
.stat-card{background:#fff;border-radius:var(--radius-sm);padding:20px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.stat-value{font-size:28px;font-weight:700;color:var(--primary)}
.stat-value.red{color:#DC2626}
.stat-value.orange{color:#D97706}
.stat-label{font-size:13px;color:var(--gray-500);margin-top:4px}
.category-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:12px}
.cat-card{background:var(--gray-50);border-radius:var(--radius-sm);padding:16px;text-align:center;border:1.5px solid var(--gray-200);transition:all .2s}
.cat-card:hover{border-color:var(--primary);transform:translateY(-2px);box-shadow:0 4px 12px rgba(0,0,0,.08)}
.cat-name{font-size:13px;color:var(--gray-500);margin-bottom:6px}
.cat-amount{font-size:20px;font-weight:700;color:var(--gray-800)}
.cat-pct{font-size:11px;color:var(--gray-400);margin-top:4px}
.cat-count{font-size:11px;color:var(--gray-400);margin-top:2px}
.table-container{overflow-x:auto;border-radius:var(--radius-sm);border:1px solid var(--gray-200)}
table{width:100%;border-collapse:collapse;font-size:14px}
thead th{background:var(--primary);color:#fff;padding:12px 10px;text-align:left;font-weight:500;white-space:nowrap;position:sticky;top:0}
tbody td{padding:10px;border-bottom:1px solid var(--gray-100);white-space:nowrap}
tbody tr:hover{background:var(--primary-light)}
tbody tr.invoice-row{background:#F0FDF4}
tbody tr.supporting-row{background:#FEF3C7}
tbody tr.duplicate-row{background:#FEE2E2 !important}
tbody tr.duplicate-row:hover{background:#FECACA !important}
.badge{display:inline-block;padding:3px 10px;border-radius:99px;font-size:12px;font-weight:500}
.badge-invoice{background:#D1FAE5;color:#059669}
.badge-supporting{background:#FEF3C7;color:#D97706}
.badge-duplicate{background:#FEE2E2;color:#DC2626}
.action-row{display:flex;gap:12px;margin-bottom:24px;flex-wrap:wrap}
.email-alert{background:var(--warning-light);border:1px solid #F59E0B;border-radius:var(--radius-sm);padding:16px;margin-bottom:16px}
.email-alert h4{color:var(--warning);margin-bottom:8px}
.email-alert ul{list-style:disc;padding-left:20px;font-size:14px;color:var(--gray-700)}
.email-alert li{margin-bottom:4px}
.email-alert a{color:var(--primary);text-decoration:underline}
.steps{display:flex;justify-content:center;gap:8px;margin-bottom:32px}
.step{display:flex;align-items:center;gap:8px;padding:8px 16px;border-radius:99px;font-size:14px;font-weight:500;background:rgba(255,255,255,.15);color:rgba(255,255,255,.6);transition:all .3s}
.step.active{background:rgba(255,255,255,.25);color:#fff}
.step.done{background:rgba(255,255,255,.3);color:#fff}
.step-num{width:24px;height:24px;border-radius:50%;background:rgba(255,255,255,.2);display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700}
.step.active .step-num{background:#fff;color:var(--primary)}
.step.done .step-num{background:#10B981;color:#fff}
.step-arrow{color:rgba(255,255,255,.3)}
.footer{text-align:center;padding:32px 0 16px;color:rgba(255,255,255,.5);font-size:13px}
.toast{position:fixed;top:20px;right:20px;padding:14px 20px;border-radius:var(--radius-sm);color:#fff;font-weight:500;font-size:14px;z-index:9999;transform:translateX(120%);transition:transform .3s ease;max-width:360px}
.toast.show{transform:translateX(0)}
.toast-success{background:var(--success)}
.toast-error{background:var(--danger)}
.toast-info{background:var(--primary)}
.security-note{background:var(--primary-light);border-radius:var(--radius-sm);padding:14px 16px;margin-top:16px;display:flex;gap:10px;align-items:flex-start;font-size:13px;color:var(--gray-600)}
.security-note .lock-icon{font-size:18px;flex-shrink:0;margin-top:1px}
@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
.fade-in{animation:fadeIn .4s ease}
.legend{display:flex;gap:16px;margin-bottom:16px;font-size:13px;color:var(--gray-500)}
.legend-item{display:flex;align-items:center;gap:6px}
.legend-dot{width:12px;height:12px;border-radius:3px}
.legend-dot-green{background:#D1FAE5;border:1px solid #059669}
.legend-dot-yellow{background:#FEF3C7;border:1px solid #D97706}
.legend-dot-red{background:#FEE2E2;border:1px solid #DC2626}
</style>
</head>
<body>
<div class="container">
<div class="header">
<h1>🧾 发票收集器</h1>
<p>输入邮箱信息，一键收集全部发票并生成统计表格</p>
</div>
<div class="steps">
<div class="step active" id="step1"><span class="step-num">1</span>配置邮箱</div>
<span class="step-arrow">→</span>
<div class="step" id="step2"><span class="step-num">2</span>收集发票</div>
<span class="step-arrow">→</span>
<div class="step" id="step3"><span class="step-num">3</span>查看结果</div>
</div>
<div class="card" id="configCard">
<div class="card-title"><span class="icon icon-blue">📧</span>邮箱配置</div>
<form id="configForm" onsubmit="startCollection(event)">
<div class="form-grid">
<div class="form-group">
<label class="form-label">邮箱地址</label>
<input type="email" class="form-input" id="email" placeholder="yourname@163.com" required>
<span class="form-hint">支持163、QQ、Gmail等支持IMAP的邮箱</span>
</div>
<div class="form-group">
<label class="form-label">IMAP授权码</label>
<input type="password" class="form-input" id="authCode" placeholder="输入IMAP授权码" required>
<span class="form-hint">非登录密码，需在邮箱设置中开启IMAP并获取授权码</span>
</div>
<div class="form-group">
<label class="form-label">IMAP服务器</label>
<select class="form-select" id="imapHost" onchange="updatePort()">
<option value="imap.163.com">163邮箱</option>
<option value="imap.qq.com">QQ邮箱</option>
<option value="imap.gmail.com">Gmail</option>
<option value="imap.outlook.com">Outlook</option>
<option value="imap.126.com">126邮箱</option>
<option value="imap.sina.com">新浪邮箱</option>
<option value="custom">自定义服务器</option>
</select>
</div>
<div class="form-group" id="customHostGroup" style="display:none">
<label class="form-label">自定义IMAP服务器</label>
<input type="text" class="form-input" id="customHost" placeholder="imap.example.com">
</div>
<div class="form-group">
<label class="form-label">开始日期</label>
<input type="date" class="form-input" id="dateFrom" required>
</div>
<div class="form-group">
<label class="form-label">结束日期</label>
<input type="date" class="form-input" id="dateTo" required>
</div>
</div>
<div class="security-note">
<span class="lock-icon">🔒</span>
<span>您的邮箱授权码仅用于本次IMAP连接读取邮件，不会被存储或上传。所有处理均在本地完成。</span>
</div>
<div style="text-align:center;margin-top:28px;">
<button type="submit" class="btn btn-primary" id="startBtn">🚀 开始收集发票</button>
</div>
</form>
</div>
<div class="card" id="progressCard" style="display:none">
<div class="card-title"><span class="icon icon-blue">⏳</span>正在收集发票</div>
<div class="progress-bar-container"><div class="progress-bar" id="progressBar"></div></div>
<div class="progress-text" id="progressText">正在连接邮箱...</div>
<div style="text-align:center;margin-top:20px;"><div class="spinner"></div></div>
</div>
<div id="resultSection" style="display:none">
<div class="stats-grid fade-in">
<div class="stat-card"><div class="stat-value" id="statEmails">0</div><div class="stat-label">扫描邮件</div></div>
<div class="stat-card"><div class="stat-value" id="statInvoices">0</div><div class="stat-label">发票数量</div></div>
<div class="stat-card"><div class="stat-value orange" id="statDuplicates">0</div><div class="stat-label">重复发票</div></div>
<div class="stat-card"><div class="stat-value" id="statSupporting">0</div><div class="stat-label">辅助文件</div></div>
<div class="stat-card"><div class="stat-value red" id="statAmount">¥0</div><div class="stat-label">发票总额</div></div>
</div>
<div id="categorySection" style="display:none" class="card fade-in">
<div class="card-title"><span class="icon icon-green">💰</span>分类金额汇总</div>
<div class="category-grid" id="categoryGrid"></div>
</div>
<div class="action-row fade-in">
<button class="btn btn-success" onclick="downloadExcel()">📥 下载Excel表格</button>
<button class="btn btn-primary" onclick="downloadAllFiles()">📦 下载全部文件(ZIP)</button>
<button class="btn btn-outline" onclick="newCollection()">🔄 新建收集任务</button>
</div>
<div id="emailAlertSection" style="display:none">
<div class="email-alert fade-in">
<h4>⚠️ 发现无附件的发票邮件</h4>
<p style="margin-bottom:8px;font-size:14px;">以下邮件包含发票信息但无附件，可能需要手动下载：</p>
<ul id="emailAlertList"></ul>
</div>
</div>
<div class="card fade-in">
<div class="card-title"><span class="icon icon-green">📊</span>发票详情</div>
<div class="legend">
<div class="legend-item"><span class="legend-dot legend-dot-green"></span>发票</div>
<div class="legend-item"><span class="legend-dot legend-dot-yellow"></span>辅助单据</div>
<div class="legend-item"><span class="legend-dot legend-dot-red"></span>重复发票</div>
</div>
<div class="table-container">
<table>
<thead><tr><th>序号</th><th>类型</th><th>地区</th><th>类目</th><th>日期</th><th>发票号码</th><th>金额</th><th>购买方</th><th>销售方</th><th>备注</th></tr></thead>
<tbody id="invoiceTableBody"></tbody>
</table>
</div>
</div>
<div class="card fade-in">
<div class="card-title"><span class="icon icon-blue">📁</span>已收集文件</div>
<div id="filesList"></div>
</div>
</div>
</div>
<div class="footer">发票收集器 v2.0 · 所有数据本地处理，安全可靠</div>
<div class="toast" id="toast"></div>
<script>
const API_BASE=window.location.origin;
let currentTaskId=null,pollInterval=null;
const today=new Date();
const threeMonthsAgo=new Date(today);
threeMonthsAgo.setMonth(threeMonthsAgo.getMonth()-3);
document.getElementById('dateFrom').value=threeMonthsAgo.toISOString().split('T')[0];
document.getElementById('dateTo').value=today.toISOString().split('T')[0];
document.getElementById('email').addEventListener('blur',function(){
const email=this.value.toLowerCase();
const select=document.getElementById('imapHost');
if(email.includes('163.com'))select.value='imap.163.com';
else if(email.includes('qq.com'))select.value='imap.qq.com';
else if(email.includes('gmail.com'))select.value='imap.gmail.com';
else if(email.includes('outlook.com')||email.includes('hotmail.com'))select.value='imap.outlook.com';
else if(email.includes('126.com'))select.value='imap.126.com';
else if(email.includes('sina.com'))select.value='imap.sina.com';
updatePort();
});
function updatePort(){
const select=document.getElementById('imapHost');
document.getElementById('customHostGroup').style.display=select.value==='custom'?'block':'none';
}
function showToast(msg,type='info'){
const toast=document.getElementById('toast');
toast.textContent=msg;
toast.className='toast toast-'+type+' show';
setTimeout(()=>toast.classList.remove('show'),4000);
}
function setStep(stepNum){
['step1','step2','step3'].forEach((id,i)=>{
const el=document.getElementById(id);
el.className='step';
if(i+1<stepNum)el.classList.add('done');
if(i+1===stepNum)el.classList.add('active');
});
}
async function startCollection(e){
e.preventDefault();
const email=document.getElementById('email').value.trim();
const authCode=document.getElementById('authCode').value.trim();
const imapHost=document.getElementById('imapHost').value==='custom'
?document.getElementById('customHost').value.trim()
:document.getElementById('imapHost').value;
const dateFrom=document.getElementById('dateFrom').value;
const dateTo=document.getElementById('dateTo').value;
if(!email||!authCode||!dateFrom||!dateTo){showToast('请填写所有必填项','error');return;}
document.getElementById('configCard').style.display='none';
document.getElementById('progressCard').style.display='block';
setStep(2);
try{
const resp=await fetch(API_BASE+'/api/collect',{
method:'POST',
headers:{'Content-Type':'application/json'},
body:JSON.stringify({email,auth_code:authCode,host:imapHost,port:993,date_from:dateFrom,date_to:dateTo})
});
if(!resp.ok)throw new Error('请求失败: '+resp.status);
const data=await resp.json();
currentTaskId=data.task_id;
pollInterval=setInterval(pollStatus,2000);
showToast('任务已创建，正在收集...','info');
}catch(err){
showToast('启动失败: '+err.message,'error');
document.getElementById('configCard').style.display='block';
document.getElementById('progressCard').style.display='none';
setStep(1);
}
}
async function pollStatus(){
if(!currentTaskId)return;
try{
const resp=await fetch(API_BASE+'/api/status/'+currentTaskId);
const data=await resp.json();
const progressMap={'pending':5,'running':50,'completed':100,'failed':100};
const pct=progressMap[data.status]||0;
document.getElementById('progressBar').style.width=pct+'%';
document.getElementById('progressText').textContent=data.progress;
if(data.status==='completed'){clearInterval(pollInterval);showResults(data.result);setStep(3);}
else if(data.status==='failed'){clearInterval(pollInterval);showToast('收集失败: '+data.progress,'error');document.getElementById('configCard').style.display='block';document.getElementById('progressCard').style.display='none';setStep(1);}
}catch(err){console.error('Poll error:',err);}
}
function showResults(result){
document.getElementById('progressCard').style.display='none';
document.getElementById('resultSection').style.display='block';
document.getElementById('statEmails').textContent=result.total_emails;
document.getElementById('statInvoices').textContent=result.invoice_count;
document.getElementById('statDuplicates').textContent=result.duplicate_count||0;
document.getElementById('statSupporting').textContent=result.supporting_count;
document.getElementById('statAmount').textContent='¥'+result.total_amount.toLocaleString('zh-CN',{minimumFractionDigits:2});
// Render category summary
if(result.category_summary&&Object.keys(result.category_summary).length>0){
document.getElementById('categorySection').style.display='block';
const grid=document.getElementById('categoryGrid');
const total=result.total_amount||0;
let html='';
const sorted=Object.entries(result.category_summary).sort((a,b)=>b[1]-a[1]);
sorted.forEach(function(item){
const cat=item[0],amt=item[1];
const pct=total>0?(amt/total*100).toFixed(1):'0';
const count=result.invoices?result.invoices.filter(function(inv){return inv.category===cat&&!inv.is_duplicate;}).length:0;
html+='<div class="cat-card"><div class="cat-name">'+cat+'</div><div class="cat-amount">¥'+amt.toLocaleString('zh-CN',{minimumFractionDigits:2})+'</div><div class="cat-pct">'+pct+'%</div><div class="cat-count">'+count+'张</div></div>';
});
grid.innerHTML=html;
}
// Render table using all_items (interleaved order)
const tbody=document.getElementById('invoiceTableBody');
tbody.innerHTML='';
if(result.all_items&&result.all_items.length>0){
result.all_items.forEach(function(item){
const tr=document.createElement('tr');
if(item.is_duplicate){
tr.className='duplicate-row';
}else if(item.item_type==='invoice'){
tr.className='invoice-row';
}else{
tr.className='supporting-row';
}
let badge='';
if(item.is_duplicate){
badge='<span class="badge badge-duplicate">重复</span>';
}else if(item.item_type==='invoice'){
badge='<span class="badge badge-invoice">发票</span>';
}else{
badge='<span class="badge badge-supporting">辅助</span>';
}
let amountCell='-';
if(item.amount){
amountCell='<span style="color:#DC2626;font-weight:600;">¥'+item.amount+'</span>';
}
tr.innerHTML='<td>'+item.seq+'</td><td>'+badge+'</td><td>'+(item.region||'-')+'</td><td>'+(item.category||'-')+'</td><td>'+(item.date||'-')+'</td><td style="font-family:monospace;font-size:12px;">'+(item.invoice_no||'-')+'</td><td>'+amountCell+'</td><td>'+(item.buyer||'-')+'</td><td>'+(item.seller||'-')+'</td><td style="font-size:12px;'+(item.is_duplicate?'color:#DC2626;font-weight:500;':'')+'">'+(item.remark||'')+'</td>';
tbody.appendChild(tr);
});
}else{
// Fallback: use separate invoices and supporting_docs
result.invoices.forEach(function(inv){
const tr=document.createElement('tr');
tr.className=inv.is_duplicate?'duplicate-row':'invoice-row';
let badge=inv.is_duplicate?'<span class="badge badge-duplicate">重复</span>':'<span class="badge badge-invoice">发票</span>';
tr.innerHTML='<td>'+inv.seq+'</td><td>'+badge+'</td><td>'+inv.region+'</td><td>'+inv.category+'</td><td>'+inv.date+'</td><td style="font-family:monospace;font-size:12px;">'+inv.invoice_no+'</td><td style="color:#DC2626;font-weight:600;">¥'+inv.amount+'</td><td>'+inv.buyer+'</td><td>'+inv.seller+'</td><td style="font-size:12px;color:#DC2626;">'+(inv.remark||'')+'</td>';
tbody.appendChild(tr);
});
result.supporting_docs.forEach(function(sd){
const tr=document.createElement('tr');
tr.className='supporting-row';
tr.innerHTML='<td>'+sd.seq+'</td><td><span class="badge badge-supporting">辅助</span></td><td>-</td><td>'+sd.type+'</td><td>-</td><td>-</td><td>-</td><td>-</td><td>-</td><td></td>';
tbody.appendChild(tr);
});
}
if(result.no_attach_invoice_emails&&result.no_attach_invoice_emails.length>0){
document.getElementById('emailAlertSection').style.display='block';
const ul=document.getElementById('emailAlertList');
ul.innerHTML='';
result.no_attach_invoice_emails.forEach(function(e){
const li=document.createElement('li');
let html='<strong>'+e.subject+'</strong> (来自: '+e.from+')';
if(e.invoice_urls&&e.invoice_urls.length>0){
html+='<br>发票链接: '+e.invoice_urls.map(function(u){return '<a href="'+u+'" target="_blank" rel="noopener">下载链接</a>';}).join(' | ');
}
li.innerHTML=html;
ul.appendChild(li);
});
}
loadFilesList();
var msg='收集完成！共'+result.invoice_count+'张发票';
if(result.duplicate_count>0)msg+='（含'+result.duplicate_count+'张重复）';
showToast(msg,'success');
}
async function loadFilesList(){
if(!currentTaskId)return;
try{
const resp=await fetch(API_BASE+'/api/files/'+currentTaskId);
const data=await resp.json();
const container=document.getElementById('filesList');
if(!data.files||data.files.length===0){container.innerHTML='<p style="color:var(--gray-400);">暂无文件</p>';return;}
let html='<div style="display:grid;grid-template-columns:1fr auto;gap:4px 12px;font-size:13px;">';
data.files.forEach(function(f){
var size=f.size>1024*1024?(f.size/1024/1024).toFixed(1)+' MB':(f.size/1024).toFixed(0)+' KB';
html+='<div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="'+f.name+'">'+f.name+'</div><div style="color:var(--gray-400);white-space:nowrap;">'+size+'</div>';
});
html+='</div>';
container.innerHTML=html;
}catch(err){console.error('Load files error:',err);}
}
function downloadExcel(){
if(!currentTaskId)return;
window.open(API_BASE+'/api/download/'+currentTaskId,'_blank');
}
function downloadAllFiles(){
if(!currentTaskId)return;
showToast('正在打包文件，请稍候...','info');
window.open(API_BASE+'/api/download-all/'+currentTaskId,'_blank');
}
function newCollection(){
currentTaskId=null;
if(pollInterval)clearInterval(pollInterval);
document.getElementById('configCard').style.display='block';
document.getElementById('progressCard').style.display='none';
document.getElementById('resultSection').style.display='none';
document.getElementById('categorySection').style.display='none';
document.getElementById('emailAlertSection').style.display='none';
document.getElementById('progressBar').style.width='0%';
setStep(1);
}
</script>
</body>
</html>"""

@app.get("/")
async def root():
    return HTMLResponse(content=HTML_CONTENT)

if __name__ == '__main__':
    import uvicorn
    print("=" * 50)
    print("  发票收集器 v2.0 已启动")
    print("  请在浏览器打开: http://localhost:8000")
    print("  按 Ctrl+C 停止服务")
    print("=" * 50)
    uvicorn.run(app, host='0.0.0.0', port=8000)
