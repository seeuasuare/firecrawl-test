#!/usr/bin/env python3
"""
Unified scraper + auditor for "1000 + 86 клиник.xlsx".

Both sheets have identical columns:
  1: Name  2: URL  3: Description  4: Emails  5: Phones
  6: Сайт работает?  7: Тематика  8: Уверенность %
  9: Детали тематики  10: Активность  11: Детали активности

For each row (skipping already filled cells):
  - Scrapes DESCRIPTION (meta description / og:description / first paragraph)
  - Scrapes EMAILS from main + contact pages
  - Scrapes PHONES from main + contact pages
  - Runs AUDIT: accessibility, theme relevance, recent activity

Usage:
    pip install openpyxl aiohttp beautifulsoup4
    python scrape_all.py

Input:  1000 + 86 клиник.xlsx
Output: 1000 + 86 клиник_result.xlsx
"""

import asyncio
import html as html_module
import re
import time
import warnings
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlparse

import aiohttp
import openpyxl
from openpyxl.styles import Font, PatternFill
from bs4 import BeautifulSoup, XMLParsedAsHTMLWarning

warnings.filterwarnings("ignore", category=XMLParsedAsHTMLWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)

# ===================== Configuration =====================

CONCURRENCY = 25
TIMEOUT = 12
SITE_TIMEOUT = 50
MONTHS_BACK = 3

INPUT_FILE = '1000 + 86 клиник.xlsx'
OUTPUT_FILE = '1000 + 86 клиник_result.xlsx'

# Column numbers (same for both sheets)
COL_NAME = 1
COL_URL = 2
COL_DESC = 3
COL_EMAIL = 4
COL_PHONE = 5
COL_STATUS = 6
COL_THEME = 7
COL_CONFIDENCE = 8
COL_THEME_DETAIL = 9
COL_ACTIVITY = 10
COL_ACTIVITY_DETAIL = 11

GREEN = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
RED = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
YELLOW = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')

CONTACT_PATHS = [
    '/contacts', '/contact', '/kontakty', '/about',
    '/about-us', '/o-nas', '/o-kompanii', '/kontakt',
    '/contact-us', '/svyaz', '/feedback', '/team',
    '/specialists', '/specialisty', '/komanda',
    '/o-centre', '/o-klinike', '/o-tsentre',
    '/rekvizity', '/politika-konfidencialnosti',
]


# ===================== Description extraction =====================

def extract_description(html):
    """Extract site description from meta tags or first paragraph."""
    soup = BeautifulSoup(html, 'html.parser')

    # 1) og:description
    og = soup.find('meta', property='og:description')
    if og and og.get('content', '').strip():
        return og['content'].strip()[:500]

    # 2) meta description
    meta = soup.find('meta', attrs={'name': 'description'})
    if meta and meta.get('content', '').strip():
        return meta['content'].strip()[:500]

    # 3) First meaningful <p> tag
    for p in soup.find_all('p'):
        text = p.get_text(strip=True)
        if len(text) > 50:
            return text[:500]

    # 4) Title tag
    title = soup.find('title')
    if title and title.get_text(strip=True):
        return title.get_text(strip=True)[:200]

    return None


# ===================== Email extraction =====================

EMAIL_RE = re.compile(
    r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}',
    re.IGNORECASE,
)

JUNK_EMAIL_DOMAINS = {
    'example.com', 'email.com', 'domain.com', 'test.com',
    'sentry.io', 'wixpress.com', 'yoursite.com', 'yourdomain.com',
    'site.com', 'w3.org', 'schema.org', 'gravatar.com', 'wordpress.org',
    'googleusercontent.com', 'tinymce.com', 'jsdelivr.net',
    'cloudflare.com', 'googleapis.com', 'gstatic.com',
    'wordpress.com', 'wp.com', 'jquery.com', 'co.com',
    'thismywebsite.com', 'domain.ru',
}

JUNK_EMAIL_PREFIXES = (
    'noreply', 'no-reply', 'support@wix', 'support@tilda',
    'webpack', 'grunt', 'gulp', 'postmaster', 'mailer-daemon',
    'admin@wordpress', 'info@starter', 'mycompany@', 'sample@',
    'privacy@', 'coolwanglu@',
)

JUNK_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '.css', '.js', '.woff'}


def is_valid_email(email):
    email = email.lower().strip()
    if len(email) > 100 or len(email) < 5 or '@' not in email:
        return False
    local, domain = email.rsplit('@', 1)
    if domain in JUNK_EMAIL_DOMAINS:
        return False
    if email.startswith(JUNK_EMAIL_PREFIXES):
        return False
    for ext in JUNK_EXTENSIONS:
        if email.endswith(ext) or ext in domain:
            return False
    if re.search(r'\d{5,}', email) or '..' in email:
        return False
    if not re.match(r'^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$', email):
        return False
    return True


def extract_emails(html):
    emails = set()
    decoded = html_module.unescape(html)
    soup = BeautifulSoup(decoded, 'html.parser')

    for a in soup.find_all('a', href=True):
        if 'mailto:' in a['href']:
            raw = a['href'].split('mailto:')[1].split('?')[0].split('#')[0].strip()
            if is_valid_email(raw):
                emails.add(raw.lower())

    for m in EMAIL_RE.findall(decoded):
        if is_valid_email(m):
            emails.add(m.lower())

    for tag in soup.find_all(True):
        for v in tag.attrs.values():
            if isinstance(v, str) and '@' in v:
                for m in EMAIL_RE.findall(v):
                    if is_valid_email(m):
                        emails.add(m.lower())

    for script in soup.find_all('script'):
        if script.string:
            for m in EMAIL_RE.findall(script.string):
                if is_valid_email(m):
                    emails.add(m.lower())

    for script in soup.find_all('script', type='application/ld+json'):
        if script.string:
            for m in EMAIL_RE.findall(script.string):
                if is_valid_email(m):
                    emails.add(m.lower())

    for meta in soup.find_all('meta'):
        c = meta.get('content', '')
        if '@' in c:
            for m in EMAIL_RE.findall(c):
                if is_valid_email(m):
                    emails.add(m.lower())

    return emails


# ===================== Phone extraction =====================

PHONE_RE = re.compile(r'(?:\+7|8)[\s\-\(]*(?:\d[\s\-\)]*){10}')
PHONE_CLEAN_RE = re.compile(r'[^\d+]')


def format_phone(digits):
    digits = digits.lstrip('+')
    if digits.startswith('8') and len(digits) == 11:
        digits = '7' + digits[1:]
    if digits.startswith('7') and len(digits) == 11:
        return f'+7{digits[1:]}'
    return f'+{digits}'


def extract_phones(html):
    phones = set()
    soup = BeautifulSoup(html, 'html.parser')

    for a in soup.find_all('a', href=True):
        if 'tel:' in a['href']:
            raw = a['href'].split('tel:')[1].split('?')[0].strip()
            cleaned = PHONE_CLEAN_RE.sub('', raw)
            if len(cleaned) >= 11:
                phones.add(format_phone(cleaned))

    text = soup.get_text(' ', strip=True)
    for m in PHONE_RE.finditer(text):
        cleaned = PHONE_CLEAN_RE.sub('', m.group(0))
        if len(cleaned) >= 11:
            phones.add(format_phone(cleaned))

    return phones


# ===================== Theme checking =====================

PRIMARY_KEYWORDS = [
    'детский психолог', 'детская психолог', 'детских психолог',
    'нейропсихолог', 'нейро-психолог',
    'детский центр', 'детский развива',
    'логопед', 'дефектолог',
    'детский нейро', 'нейрокоррекц',
    'психолог для детей', 'психолог для ребенк', 'психолог для ребёнк',
    'психологическ центр', 'психологический центр',
    'коррекционн', 'коррекция',
    'развити ребенк', 'развити ребёнк', 'развитие детей',
    'aba-терап', 'аба-терап', 'аба терап',
    'сенсорная интеграц', 'канистерап',
    'детский невролог', 'детский психиатр',
    'песочная терап', 'арт-терап', 'игровая терап',
    'задержк развит', 'зпр', 'зрр', 'сдвг', 'аутизм', 'рас ',
    'дислекси', 'дисграф', 'алали', 'заикани',
]

SECONDARY_KEYWORDS = [
    'психолог', 'психотерап', 'консультац', 'тревожност', 'эмоциональн',
]

ANTI_KEYWORDS = [
    'автосервис', 'автомобил', 'ремонт квартир', 'натяжн потолк',
    'юридическ', 'бухгалтер', 'строительств', 'недвижимост',
    'стоматолог', 'кулинар', 'доставка еды',
]


def check_theme(html):
    if not html:
        return None, 0, 'Нет данных'
    text = BeautifulSoup(html, 'html.parser').get_text(' ', strip=True).lower()
    if sum(1 for kw in ANTI_KEYWORDS if kw in text) >= 2:
        return False, 0, 'Другая тематика'
    pf = [kw for kw in PRIMARY_KEYWORDS if kw in text]
    sf = [kw for kw in SECONDARY_KEYWORDS if kw in text]
    if len(pf) >= 3:
        return True, 100, f'Точное совпадение ({len(pf)} ключевых слов)'
    if len(pf) >= 1:
        c = min(50 + len(pf) * 20 + len(sf) * 10, 100)
        return True, c, f'Совпадение ({len(pf)} осн. + {len(sf)} доп.)'
    if len(sf) >= 2:
        return True, 40, f'Возможное совпадение ({len(sf)} доп. слов)'
    if len(sf) == 1:
        return None, 20, 'Слабое совпадение (1 общее слово)'
    return False, 0, 'Не соответствует тематике'


# ===================== Activity checking =====================

RUSSIAN_MONTHS = {
    'январ': 1, 'феврал': 2, 'март': 3, 'апрел': 4,
    'мая': 5, 'мае': 5, 'май': 5, 'июн': 6, 'июл': 7,
    'август': 8, 'сентябр': 9, 'октябр': 10, 'ноябр': 11, 'декабр': 12,
}

DATE_PATTERNS = [
    re.compile(r'(\d{1,2})\s+(' + '|'.join(RUSSIAN_MONTHS.keys()) + r')[а-яё]*\s+(\d{4})', re.I),
    re.compile(r'(\d{1,2})[./](\d{1,2})[./](\d{4})'),
    re.compile(r'(\d{4})-(\d{2})-(\d{2})'),
]

COPYRIGHT_RE = re.compile(r'©\s*(?:\d{4}\s*[-–—]\s*)?(\d{4})', re.I)


def parse_date(match, idx):
    try:
        if idx == 0:
            d, mt, y = int(match.group(1)), match.group(2).lower(), int(match.group(3))
            mo = next((v for k, v in RUSSIAN_MONTHS.items() if mt.startswith(k)), None)
            if mo and 2000 <= y <= 2030:
                return datetime(y, mo, d)
        elif idx == 1:
            d, mo, y = int(match.group(1)), int(match.group(2)), int(match.group(3))
            if 2000 <= y <= 2030 and 1 <= mo <= 12:
                return datetime(y, mo, d)
        elif idx == 2:
            y, mo, d = int(match.group(1)), int(match.group(2)), int(match.group(3))
            if 2000 <= y <= 2030 and 1 <= mo <= 12:
                return datetime(y, mo, d)
    except (ValueError, IndexError):
        pass
    return None


def check_activity(html, headers):
    now = datetime.now()
    cutoff = now - timedelta(days=MONTHS_BACK * 30)
    results = []

    lm = headers.get('Last-Modified', '')
    if lm:
        try:
            from email.utils import parsedate_to_datetime
            lm_dt = parsedate_to_datetime(lm).replace(tzinfo=None)
            if lm_dt >= cutoff:
                results.append(f'Last-Modified: {lm_dt.strftime("%d.%m.%Y")}')
        except Exception:
            pass

    if not html:
        return (True, '; '.join(results)) if results else (None, 'Нет данных')

    text = BeautifulSoup(html, 'html.parser').get_text(' ', strip=True)
    dates = []
    for i, pat in enumerate(DATE_PATTERNS):
        for m in pat.finditer(text):
            dt = parse_date(m, i)
            if dt:
                dates.append(dt)

    recent = [d for d in dates if cutoff <= d <= now + timedelta(days=30)]
    if recent:
        results.append(f'Дата на сайте: {max(recent).strftime("%d.%m.%Y")} ({len(recent)} свежих)')

    cr = COPYRIGHT_RE.search(text)
    if cr:
        yr = int(cr.group(1))
        if yr >= now.year:
            results.append(f'Copyright {yr}')
        elif yr == now.year - 1:
            results.append(f'Copyright {yr} (прошлый год)')

    if results:
        return True, '; '.join(results)
    if dates:
        return False, f'Последняя дата: {max(dates).strftime("%d.%m.%Y")}'
    if cr and int(cr.group(1)) < now.year - 1:
        return False, f'Copyright {int(cr.group(1))} (устаревший)'
    return None, 'Нет признаков активности'


# ===================== Fetching =====================

async def fetch_page(session, url):
    try:
        async with session.get(
            url, timeout=aiohttp.ClientTimeout(total=TIMEOUT),
            ssl=False, allow_redirects=True,
        ) as resp:
            if resp.status == 200:
                ct = resp.headers.get('Content-Type', '')
                if 'text' in ct or 'html' in ct or not ct:
                    return resp.status, await resp.text(errors='replace'), dict(resp.headers)
            return resp.status, '', dict(resp.headers)
    except aiohttp.ClientConnectorError:
        return 0, '', {}
    except asyncio.TimeoutError:
        return -1, '', {}
    except Exception:
        return -2, '', {}


async def fetch_contact_pages(session, base_url, main_html, netloc):
    extra = []
    checked = set()

    if main_html:
        soup = BeautifulSoup(main_html, 'html.parser')
        kws = ['contact', 'kontakt', 'контакт', 'связ', 'о нас', 'about', 'обратн', 'напис']
        for a in soup.find_all('a', href=True):
            hl = a['href'].lower()
            tl = (a.get_text() or '').lower()
            if any(k in hl or k in tl for k in kws):
                full = urljoin(base_url, a['href'])
                if urlparse(full).netloc == netloc and full not in checked:
                    checked.add(full)
                    _, h, _ = await fetch_page(session, full)
                    if h:
                        extra.append(h)
                    if len(checked) >= 5:
                        break

    for path in CONTACT_PATHS[:10]:
        cu = base_url + path
        if cu not in checked:
            checked.add(cu)
            _, h, _ = await fetch_page(session, cu)
            if h:
                extra.append(h)

    return extra


def status_text(code):
    if code == 200: return 'OK'
    if code == 0: return 'Не открывается (ошибка соединения)'
    if code == -1: return 'Не открывается (таймаут)'
    if code == -2: return 'Не открывается (ошибка)'
    if code == 403: return 'Заблокирован (403)'
    if code == 404: return 'Не найден (404)'
    if 500 <= code < 600: return f'Ошибка сервера ({code})'
    if 300 <= code < 400: return f'Редирект ({code})'
    return f'Ошибка ({code})'


# ===================== Process one row =====================

def cell_empty(ws, row, col):
    v = ws.cell(row, col).value
    return not v or not str(v).strip()


async def process_row(session, url, needs):
    """needs = dict with keys: desc, email, phone, audit (booleans)"""
    result = {}

    if not url or url.strip() in ('', 'None'):
        return result

    url = url.strip()
    if not url.startswith('http'):
        url = 'https://' + url

    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    netloc = parsed.netloc

    try:
        async with asyncio.timeout(SITE_TIMEOUT):
            status_code, main_html, headers = await fetch_page(session, url)

            all_html = [main_html] if main_html else []

            # Fetch contact pages if needed
            if main_html and (needs.get('email') or needs.get('phone')):
                extra = await fetch_contact_pages(session, base, main_html, netloc)
                all_html.extend(extra)

            # Description
            if needs.get('desc') and main_html:
                desc = extract_description(main_html)
                if desc:
                    result['desc'] = desc

            # Emails
            if needs.get('email'):
                emails = set()
                for h in all_html:
                    emails.update(extract_emails(h))
                if emails:
                    result['emails'] = ', '.join(sorted(emails))

            # Phones
            if needs.get('phone'):
                phones = set()
                for h in all_html:
                    phones.update(extract_phones(h))
                if phones:
                    result['phones'] = ', '.join(sorted(phones))

            # Audit
            if needs.get('audit'):
                st = status_text(status_code)
                theme, conf, theme_det = check_theme(main_html)
                active, active_det = check_activity(main_html, headers)
                result['audit'] = (st, status_code, theme, conf, theme_det, active, active_det)

    except (asyncio.TimeoutError, TimeoutError):
        if needs.get('audit'):
            result['audit'] = ('Не открывается (таймаут)', -1, None, 0, 'Нет данных', None, 'Нет данных')

    return result


# ===================== Main =====================

async def process_sheet(session, ws, sheet_name):
    # Find actual rows (skip empty rows)
    rows_data = []
    for r in range(2, ws.max_row + 1):
        url = ws.cell(r, COL_URL).value
        if not url or not str(url).strip():
            continue
        rows_data.append(r)

    total = len(rows_data)
    print(f"\n  Sheet '{sheet_name}': {total} companies with URLs")

    # Determine what each row needs
    tasks_list = []
    for r in rows_data:
        needs = {
            'desc': cell_empty(ws, r, COL_DESC),
            'email': cell_empty(ws, r, COL_EMAIL),
            'phone': cell_empty(ws, r, COL_PHONE),
            'audit': cell_empty(ws, r, COL_STATUS),
        }
        if any(needs.values()):
            tasks_list.append((r, str(ws.cell(r, COL_URL).value).strip(), needs))

    to_process = len(tasks_list)
    need_desc = sum(1 for _, _, n in tasks_list if n['desc'])
    need_email = sum(1 for _, _, n in tasks_list if n['email'])
    need_phone = sum(1 for _, _, n in tasks_list if n['phone'])
    need_audit = sum(1 for _, _, n in tasks_list if n['audit'])

    print(f"  To process: {to_process} rows")
    print(f"    Descriptions: {need_desc}")
    print(f"    Emails:       {need_email}")
    print(f"    Phones:       {need_phone}")
    print(f"    Audit:        {need_audit}")

    if to_process == 0:
        print("  Nothing to do!")
        return

    processed = 0
    counts = {'desc': 0, 'email': 0, 'phone': 0}

    sem = asyncio.Semaphore(CONCURRENCY)

    async def do_row(row_idx, url, needs):
        nonlocal processed
        async with sem:
            result = await process_row(session, url, needs)

            if 'desc' in result:
                ws.cell(row_idx, COL_DESC).value = result['desc']
                counts['desc'] += 1

            if 'emails' in result:
                ws.cell(row_idx, COL_EMAIL).value = result['emails']
                counts['email'] += 1

            if 'phones' in result:
                ws.cell(row_idx, COL_PHONE).value = result['phones']
                counts['phone'] += 1

            if 'audit' in result:
                st, st_code, theme, conf, theme_det, active, active_det = result['audit']

                ws.cell(row_idx, COL_STATUS).value = st
                if st_code == 200:
                    ws.cell(row_idx, COL_STATUS).fill = GREEN
                elif st_code > 0:
                    ws.cell(row_idx, COL_STATUS).fill = YELLOW
                else:
                    ws.cell(row_idx, COL_STATUS).fill = RED

                if theme is True:
                    ws.cell(row_idx, COL_THEME).value = 'Да'
                    ws.cell(row_idx, COL_THEME).fill = GREEN
                elif theme is False:
                    ws.cell(row_idx, COL_THEME).value = 'Нет'
                    ws.cell(row_idx, COL_THEME).fill = RED
                else:
                    ws.cell(row_idx, COL_THEME).value = 'Неясно'
                    ws.cell(row_idx, COL_THEME).fill = YELLOW

                ws.cell(row_idx, COL_CONFIDENCE).value = conf
                ws.cell(row_idx, COL_THEME_DETAIL).value = theme_det

                if active is True:
                    ws.cell(row_idx, COL_ACTIVITY).value = 'Активен'
                    ws.cell(row_idx, COL_ACTIVITY).fill = GREEN
                elif active is False:
                    ws.cell(row_idx, COL_ACTIVITY).value = 'Неактивен'
                    ws.cell(row_idx, COL_ACTIVITY).fill = RED
                else:
                    ws.cell(row_idx, COL_ACTIVITY).value = 'Неизвестно'
                    ws.cell(row_idx, COL_ACTIVITY).fill = YELLOW

                ws.cell(row_idx, COL_ACTIVITY_DETAIL).value = active_det

            processed += 1
            if processed % 15 == 0 or processed == to_process:
                pct = processed * 100 // to_process
                bar = '\u2588' * (pct // 5) + '\u2591' * (20 - pct // 5)
                print(f"\r  [{bar}] {processed}/{to_process}  "
                      f"desc:+{counts['desc']} email:+{counts['email']} "
                      f"phone:+{counts['phone']}",
                      end='', flush=True)

    tasks = [do_row(r, u, n) for r, u, n in tasks_list]
    await asyncio.gather(*tasks)

    print(f"\n  => Done! desc:+{counts['desc']}  email:+{counts['email']}  phone:+{counts['phone']}")


async def main():
    print(f"Loading {INPUT_FILE}...")
    wb = openpyxl.load_workbook(INPUT_FILE)

    connector = aiohttp.TCPConnector(limit=CONCURRENCY, ssl=False)
    headers = {
        'User-Agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/120.0.0.0 Safari/537.36'
        ),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ru-RU,ru;q=0.9,en;q=0.8',
    }

    async with aiohttp.ClientSession(connector=connector, headers=headers) as session:
        for sheet_name in wb.sheetnames:
            await process_sheet(session, wb[sheet_name], sheet_name)

    wb.save(OUTPUT_FILE)
    print(f"\n{'='*55}")
    print(f"  Results saved to: {OUTPUT_FILE}")
    print(f"{'='*55}")


if __name__ == '__main__':
    start = time.time()
    asyncio.run(main())
    elapsed = time.time() - start
    m, s = divmod(int(elapsed), 60)
    print(f"Total time: {m}m {s}s")
