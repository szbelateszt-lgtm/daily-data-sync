"""
Foodora Lead Finder - Új éttermek és élelmiszerboltok automatikus felderítése
Budapest + agglomeráció területén.

Adatforrások:
1. Google Places API (New) - hivatalos, megbízható
2. crt.sh - új .hu domainek vendéglátós kulcsszavakkal
3. Google News RSS - friss hírek új helyekről
4. We Love Budapest gasztro - frissen nyílt helyek cikkei
5. Funzine.hu gasztro - frissen nyílt helyek cikkei
6. Time Out Budapest - frissen nyílt helyek cikkei

A script naponta egyszer fut le GitHub Actions-ben.
Az új találatokat e-mailben küldi el.
"""

import os
import re
import json
import smtplib
import ssl
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
import requests
import feedparser

# ============================================================
# KONFIGURÁCIÓ
# ============================================================

GOOGLE_API_KEY = os.environ.get("GOOGLE_API_KEY", "")
EMAIL_FROM = os.environ.get("EMAIL_FROM", "")
EMAIL_TO = os.environ.get("EMAIL_TO", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")  # Gmail app password
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# Keresési területek: Budapest 23 kerülete + agglomeráció főbb városai
# Minden terület egy kör középponttal és sugárral (méterben)
SEARCH_AREAS = [
    # Budapest belváros (V., VI., VII., VIII., IX.)
    {"name": "Budapest belváros", "lat": 47.4979, "lng": 19.0521, "radius": 3000},
    # Budapest észak (II., III., IV., XIII., XV.)
    {"name": "Budapest észak", "lat": 47.5500, "lng": 19.0700, "radius": 5000},
    # Budapest dél (XI., XX., XXI., XXII., XXIII.)
    {"name": "Budapest dél", "lat": 47.4500, "lng": 19.0700, "radius": 5000},
    # Budapest kelet (X., XIV., XVI., XVII., XVIII., XIX.)
    {"name": "Budapest kelet", "lat": 47.4900, "lng": 19.1500, "radius": 5000},
    # Budapest nyugat (I., XII.)
    {"name": "Budapest nyugat", "lat": 47.5100, "lng": 19.0100, "radius": 3000},
    # Agglomeráció
    {"name": "Budaörs", "lat": 47.4626, "lng": 18.9580, "radius": 3000},
    {"name": "Érd", "lat": 47.3950, "lng": 18.9050, "radius": 4000},
    {"name": "Szentendre", "lat": 47.6669, "lng": 19.0760, "radius": 3000},
    {"name": "Gödöllő", "lat": 47.6000, "lng": 19.3550, "radius": 3000},
    {"name": "Vecsés", "lat": 47.4100, "lng": 19.2700, "radius": 3000},
    {"name": "Dunakeszi", "lat": 47.6350, "lng": 19.1400, "radius": 3000},
]

# Hely típusok, amik érdekelnek minket (Foodora-szempontból releváns)
RELEVANT_TYPES = [
    "restaurant",
    "cafe",
    "bakery",
    "bar",
    "meal_takeaway",
    "meal_delivery",
    "convenience_store",
    "grocery_store",
    "supermarket",
    "food_store",
]

# Adattárolás
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
KNOWN_PLACES_FILE = DATA_DIR / "known_places.json"
KNOWN_DOMAINS_FILE = DATA_DIR / "known_domains.json"
KNOWN_WLB_FILE = DATA_DIR / "known_wlb_articles.json"
KNOWN_FUNZINE_FILE = DATA_DIR / "known_funzine_articles.json"
KNOWN_TIMEOUT_FILE = DATA_DIR / "known_timeout_articles.json"

# ============================================================
# 1. GOOGLE PLACES API - ÚJ ÉS HAMAROSAN NYÍLÓ HELYEK
# ============================================================

def fetch_google_places():
    """
    Google Places API New - Text Search hívás minden területre.
    A FUTURE_OPENING státuszú helyek mellett a frissen megjelenteket is gyűjti.
    """
    if not GOOGLE_API_KEY:
        print("⚠️  GOOGLE_API_KEY nincs beállítva, kihagyom.")
        return []

    all_places = []
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": GOOGLE_API_KEY,
        "X-Goog-FieldMask": (
            "places.id,places.displayName,places.formattedAddress,"
            "places.businessStatus,places.openingDate,places.primaryType,"
            "places.googleMapsUri,places.websiteUri,"
            "places.nationalPhoneNumber,places.location"
        ),
    }

    for area in SEARCH_AREAS:
        # FUTURE_OPENING helyek - ezek hamarosan nyílnak, a legértékesebb leadek!
        # A 'restaurant' includedType szándékosan szűkebb keresést kényszerít.
        # Mivel a 'bakery', 'cafe', 'bar' stb. nem 'restaurant' típus, ezért
        # nem alkalmazunk includedType-ot - hagyjuk hogy a textQuery vezessen.
        for query in ["étterem", "kávézó", "élelmiszerbolt", "pékség", "bisztró"]:
            body = {
                "textQuery": f"{query} {area['name']}",
                "locationBias": {
                    "circle": {
                        "center": {
                            "latitude": area["lat"],
                            "longitude": area["lng"],
                        },
                        "radius": area["radius"],
                    }
                },
                "maxResultCount": 20,
                "includeFutureOpeningBusinesses": True,  # 2026 márciustól ez a helyes név
            }

            try:
                resp = requests.post(
                    "https://places.googleapis.com/v1/places:searchText",
                    headers=headers,
                    json=body,
                    timeout=30,
                )
                if resp.status_code != 200:
                    print(f"  ✗ Hiba {area['name']} / {query}: HTTP {resp.status_code} – {resp.text[:200]}")
                    continue
                data = resp.json()
                places = data.get("places", [])
                for p in places:
                    p["_search_area"] = area["name"]
                    p["_search_query"] = query
                all_places.extend(places)
                print(f"  ✓ {area['name']} / {query}: {len(places)} hely")
            except Exception as e:
                print(f"  ✗ Hiba {area['name']} / {query}: {e}")

    # Deduplikálás place_id alapján
    unique = {}
    for p in all_places:
        pid = p.get("id")
        if pid and pid not in unique:
            unique[pid] = p
    print(f"📍 Összesen {len(unique)} egyedi hely")
    return list(unique.values())


def filter_new_places(places, known_ids):
    """Csak azok a helyek érdekesek, amik még nem voltak ismertek."""
    new = []
    for p in places:
        pid = p.get("id")
        if pid and pid not in known_ids:
            new.append(p)
    return new


# ============================================================
# 2. crt.sh - ÚJ .hu DOMAIN REGISZTRÁCIÓK
# ============================================================

def fetch_new_domains():
    """
    Új SSL tanúsítványokat kérdez le a Certificate Transparency logból
    vendéglátós kulcsszavakkal. Minden új weboldal automatikusan kap SSL-t,
    így ez gyakorlatilag az új weboldalak listája.
    """
    keywords = [
        "etterem", "restaurant", "bisztro", "bistro", "kavezo",
        "pizza", "burger", "pekseg", "bakery", "kifozde",
        "kocsma", "bar", "elelmiszer", "grocery", "deli",
        "konyha", "kitchen", "food", "cafe", "trattoria",
    ]

    cutoff = (datetime.utcnow() - timedelta(days=2)).strftime("%Y-%m-%d")
    all_domains = set()

    headers = {
        "User-Agent": "Mozilla/5.0 (Foodora Lead Finder; +https://github.com/)"
    }
    for kw in keywords:
        url = f"https://crt.sh/?q=%25{kw}%25.hu&output=json"
        try:
            resp = requests.get(url, headers=headers, timeout=30)
            if resp.status_code != 200:
                print(f"  ⚠️  crt.sh '{kw}': HTTP {resp.status_code}")
                continue
            data = resp.json()
            for cert in data:
                # Csak az utóbbi 2 napban kiadott tanúsítványok
                if cert.get("entry_timestamp", "")[:10] < cutoff:
                    continue
                names = cert.get("name_value", "").split("\n")
                for name in names:
                    name = name.strip().lower()
                    # Csak .hu domainek, nem wildcard, nem aldomain
                    if (name.endswith(".hu")
                            and "*" not in name
                            and name.count(".") == 1):
                        all_domains.add(name)
            print(f"  ✓ '{kw}': eddig {len(all_domains)} unique domain")
        except Exception as e:
            print(f"  ✗ Hiba '{kw}': {e}")

    return list(all_domains)


# ============================================================
# 3. GOOGLE NEWS RSS - FRISS HÍREK
# ============================================================

def fetch_news():
    """Google News RSS keresés vendéglátós kulcsszavakkal."""
    queries = [
        "új étterem Budapest",
        "megnyílt Budapest étterem",
        "nyitás Budapest étterem",
        "új kávézó Budapest",
        "új bisztró Budapest",
        "új élelmiszerbolt Budapest",
    ]

    all_news = []
    # 7 napos cutoff (volt 2 napos, de túl szigorú volt és minden hírt eldobott)
    cutoff = datetime.utcnow() - timedelta(days=7)

    for q in queries:
        url = (
            f"https://news.google.com/rss/search?q={requests.utils.quote(q)}"
            f"&hl=hu&gl=HU&ceid=HU:hu"
        )
        try:
            feed = feedparser.parse(url)
            kept = 0
            for entry in feed.entries[:10]:
                # Próbáljuk megkapni a dátumot többféleképpen
                pub_dt = None
                pub_parsed = entry.get("published_parsed")
                if pub_parsed:
                    try:
                        pub_dt = datetime(*pub_parsed[:6])
                    except (TypeError, ValueError):
                        pub_dt = None

                # Ha nem tudjuk a dátumot, BEENGEDJÜK a hírt (jobb beengedni, mint elveszíteni)
                # Ha tudjuk és túl régi, kihagyjuk
                if pub_dt and pub_dt < cutoff:
                    continue

                all_news.append({
                    "title": entry.get("title", ""),
                    "link": entry.get("link", ""),
                    "published": entry.get("published", ""),
                    "source": entry.get("source", {}).get("title", "") if entry.get("source") else "",
                    "query": q,
                })
                kept += 1
            print(f"  ✓ '{q}': {len(feed.entries[:10])} talált, {kept} elfogadva")
        except Exception as e:
            print(f"  ✗ Hiba '{q}': {e}")

    # Deduplikálás link alapján
    seen = set()
    unique = []
    for n in all_news:
        if n["link"] not in seen:
            seen.add(n["link"])
            unique.append(n)
    return unique


# ============================================================
# GASZTROPORTÁLOK – HTML SCRAPE
# ============================================================

def _strip_html(text):
    """Eltávolítja a HTML tageket és tisztítja a whitespace-t."""
    # HTML entitások visszafordítása
    text = text.replace("&amp;", "&").replace("&nbsp;", " ")
    text = text.replace("&quot;", '"').replace("&#8211;", "–").replace("&#8217;", "'")
    # HTML tagek eltávolítása
    text = re.sub(r"<[^>]+>", "", text)
    # Több whitespace egybe
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _fetch_html(url, timeout=30):
    """HTML letöltése böngésző-szerű header-rel (anti-bot védelem ellen)."""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "hu-HU,hu;q=0.9,en;q=0.8",
    }
    resp = requests.get(url, headers=headers, timeout=timeout)
    resp.raise_for_status()
    return resp.text


def fetch_welovebudapest():
    """We Love Budapest gasztro rovat - friss cikkek a kezdőoldalról."""
    try:
        html = _fetch_html("https://welovebudapest.com/gasztro/")
    except Exception as e:
        print(f"  ✗ Hiba a We Love Budapest letöltésénél: {e}")
        return []

    articles = []
    # Két pattern-t használunk:
    # 1. ## [Cím](URL) markdown-szerű
    # 2. <h2><a href="URL">Cím</a></h2> HTML
    patterns = [
        re.compile(
            r'\[([^\]]{15,250})\]\((https://welovebudapest\.com/(?:cikk|toplista)/[^\)]+)\)',
            re.IGNORECASE,
        ),
        re.compile(
            r'<a[^>]+href="(https://welovebudapest\.com/(?:cikk|toplista)/[^"]+)"[^>]*>'
            r'\s*<h[2-4][^>]*>([\s\S]{15,300}?)</h[2-4]>',
            re.IGNORECASE,
        ),
    ]

    seen_urls = set()
    for pattern in patterns:
        for match in pattern.finditer(html):
            # Pattern 1: group(1)=title, group(2)=url
            # Pattern 2: group(1)=url, group(2)=title
            if pattern.pattern.startswith(r"\["):
                title = _strip_html(match.group(1))
                url = match.group(2)
            else:
                url = match.group(1)
                title = _strip_html(match.group(2))

            if url in seen_urls:
                continue
            if not title or len(title) < 10:
                continue
            seen_urls.add(url)
            articles.append({"title": title.strip(), "url": url})

    print(f"  ✓ We Love Budapest: {len(articles)} cikk találva")
    return articles


def fetch_funzine():
    """Funzine.hu gasztro rovat - friss cikkek."""
    try:
        html = _fetch_html("https://funzine.hu/category/gasztro/")
    except Exception as e:
        print(f"  ✗ Hiba a Funzine letöltésénél: {e}")
        return []

    articles = []
    # Funzine cikk URL minta: https://funzine.hu/YYYY/MM/DD/gasztro/SLUG/
    # Két pattern, akárcsak WLB-nél
    patterns = [
        re.compile(
            r'\[([^\]]{15,250})\]\((https://funzine\.hu/\d{4}/\d{2}/\d{2}/gasztro/[^\)]+)\)',
            re.IGNORECASE,
        ),
        re.compile(
            r'<a[^>]+href="(https://funzine\.hu/\d{4}/\d{2}/\d{2}/gasztro/[^"]+)"[^>]*'
            r'title="([^"]{15,300})"',
            re.IGNORECASE,
        ),
    ]

    seen_urls = set()
    for pattern in patterns:
        for match in pattern.finditer(html):
            if pattern.pattern.startswith(r"\["):
                title = _strip_html(match.group(1))
                url = match.group(2)
            else:
                url = match.group(1)
                title = _strip_html(match.group(2))

            if url in seen_urls:
                continue
            if not title or len(title) < 10:
                continue
            seen_urls.add(url)
            articles.append({"title": title.strip(), "url": url})

    print(f"  ✓ Funzine: {len(articles)} cikk találva")
    return articles


def fetch_timeout_budapest():
    """Time Out Budapest étterem cikkek."""
    # Több potenciális URL-en is megpróbáljuk
    html = None
    for url in [
        "https://www.timeout.com/hu/budapest/ettermek",
        "https://www.timeout.com/hu/budapest/etterem",
        "https://www.timeout.com/hu/budapest",
    ]:
        try:
            html = _fetch_html(url)
            break
        except Exception:
            continue

    if not html:
        print("  ✗ Hiba a Time Out Budapest letöltésénél: nem érhető el")
        return []

    articles = []
    # Time Out: étterem-kapcsolt URL-eket keresünk
    relevant_keywords = [
        "etterem", "ettermek", "kavezo", "bistro", "bisztro",
        "food", "gasztro", "reggeli", "brunch", "kavehaz",
        "pekseg", "restaurant", "bar"
    ]
    patterns = [
        re.compile(
            r'\[([^\]]{15,250})\]\((https://www\.timeout\.com/hu/budapest/[^\)]+)\)',
            re.IGNORECASE,
        ),
        re.compile(
            r'<a[^>]+href="(https://www\.timeout\.com/hu/budapest/[^"]+)"[^>]*'
            r'(?:title|aria-label)="([^"]{15,300})"',
            re.IGNORECASE,
        ),
    ]

    seen_urls = set()
    for pattern in patterns:
        for match in pattern.finditer(html):
            if pattern.pattern.startswith(r"\["):
                title = _strip_html(match.group(1))
                url = match.group(2)
            else:
                url = match.group(1)
                title = _strip_html(match.group(2))

            # Csak vendéglátós tematikájú cikkek
            if not any(kw in url.lower() for kw in relevant_keywords):
                continue
            if url in seen_urls:
                continue
            if not title or len(title) < 10:
                continue
            seen_urls.add(url)
            articles.append({"title": title.strip(), "url": url})

    print(f"  ✓ Time Out Budapest: {len(articles)} cikk találva")
    return articles


def filter_new_articles(articles, known_urls):
    """Csak az új cikkek, amik még nem voltak ismertek."""
    return [a for a in articles if a.get("url") not in known_urls]


# ============================================================
# E-MAIL FORMÁZÁS ÉS KÜLDÉS
# ============================================================

def format_email(new_places, new_domains, news,
                 new_wlb_articles, new_funzine_articles, new_timeout_articles):
    """HTML e-mail összeállítása az új találatokkal."""
    today = datetime.now().strftime("%Y-%m-%d")

    html = f"""
    <html>
    <head>
    <style>
      body {{ font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; color: #333; }}
      h1 {{ color: #d70f64; border-bottom: 3px solid #d70f64; padding-bottom: 10px; }}
      h2 {{ color: #d70f64; margin-top: 30px; }}
      .place {{ background: #f9f9f9; padding: 12px; margin: 8px 0; border-left: 4px solid #d70f64; border-radius: 4px; }}
      .place.future {{ border-left-color: #ff9500; background: #fff8e6; }}
      .name {{ font-weight: bold; font-size: 16px; }}
      .badge {{ display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 11px; font-weight: bold; }}
      .badge.future {{ background: #ff9500; color: white; }}
      .badge.new {{ background: #d70f64; color: white; }}
      .meta {{ color: #666; font-size: 13px; margin-top: 4px; }}
      a {{ color: #d70f64; text-decoration: none; }}
      .stats {{ background: #f0f0f0; padding: 10px; border-radius: 5px; }}
      .empty {{ color: #999; font-style: italic; }}
    </style>
    </head>
    <body>
    <h1>🍕 Foodora Lead Finder – {today}</h1>
    <div class="stats">
      <strong>Mai találatok:</strong>
      Google Maps: {len(new_places)} új hely &nbsp;|&nbsp;
      Új domainek: {len(new_domains)} &nbsp;|&nbsp;
      Hírek: {len(news)} &nbsp;|&nbsp;
      WLB: {len(new_wlb_articles)} &nbsp;|&nbsp;
      Funzine: {len(new_funzine_articles)} &nbsp;|&nbsp;
      Time Out: {len(new_timeout_articles)}
    </div>
    """

    # 1. ÚJ HELYEK GOOGLE MAPS-RŐL
    html += "<h2>📍 Új helyek a Google Maps-en</h2>"
    if not new_places:
        html += '<p class="empty">Ma nem találtunk új helyet.</p>'
    else:
        # Először a future opening, mert ezek a legértékesebbek
        future = [p for p in new_places if p.get("businessStatus") == "FUTURE_OPENING"]
        current = [p for p in new_places if p.get("businessStatus") != "FUTURE_OPENING"]

        if future:
            html += "<h3>🔥 Hamarosan nyíló helyek (TOP PRIORITÁS!)</h3>"
            for p in future:
                html += format_place_html(p, is_future=True)

        if current:
            html += "<h3>✨ Frissen megjelent helyek</h3>"
            for p in current[:50]:  # max 50, nehogy szétessen az e-mail
                html += format_place_html(p, is_future=False)

    # 2. ÚJ DOMAINEK
    html += "<h2>🌐 Új weboldalak (vendéglátós kulcsszavakkal)</h2>"
    if not new_domains:
        html += '<p class="empty">Ma nem találtunk új domaint.</p>'
    else:
        html += "<ul>"
        for d in sorted(new_domains)[:50]:
            html += f'<li><a href="https://{d}">{d}</a></li>'
        html += "</ul>"
        if len(new_domains) > 50:
            html += f"<p><em>... és még {len(new_domains)-50} domain</em></p>"

    # 3. HÍREK
    html += "<h2>📰 Friss hírek</h2>"
    if not news:
        html += '<p class="empty">Ma nem volt releváns hír.</p>'
    else:
        for n in news[:20]:
            html += f"""
            <div class="place">
              <div class="name"><a href="{n['link']}">{n['title']}</a></div>
              <div class="meta">{n.get('source', '')} • {n.get('published', '')}</div>
            </div>
            """

    # 4. WE LOVE BUDAPEST CIKKEK
    html += "<h2>🌶️ We Love Budapest – új cikkek a gasztro rovatból</h2>"
    if not new_wlb_articles:
        html += '<p class="empty">Ma nem volt új cikk.</p>'
    else:
        for a in new_wlb_articles[:25]:
            html += f"""
            <div class="place">
              <div class="name"><a href="{a['url']}">{a['title']}</a></div>
              <div class="meta">welovebudapest.com</div>
            </div>
            """

    # 5. FUNZINE CIKKEK
    html += "<h2>🎉 Funzine.hu – új cikkek a gasztro rovatból</h2>"
    if not new_funzine_articles:
        html += '<p class="empty">Ma nem volt új cikk.</p>'
    else:
        for a in new_funzine_articles[:25]:
            html += f"""
            <div class="place">
              <div class="name"><a href="{a['url']}">{a['title']}</a></div>
              <div class="meta">funzine.hu</div>
            </div>
            """

    # 6. TIME OUT BUDAPEST CIKKEK
    html += "<h2>🌍 Time Out Budapest – új éttermes cikkek</h2>"
    if not new_timeout_articles:
        html += '<p class="empty">Ma nem volt új cikk.</p>'
    else:
        for a in new_timeout_articles[:25]:
            html += f"""
            <div class="place">
              <div class="name"><a href="{a['url']}">{a['title']}</a></div>
              <div class="meta">timeout.com</div>
            </div>
            """

    html += """
    <hr>
    <p style="color: #999; font-size: 11px;">
      Ezt az e-mailt a Foodora Lead Finder automata küldte.
      Adatforrások: Google Places API, crt.sh, Google News, We Love Budapest, Funzine, Time Out Budapest.
    </p>
    </body>
    </html>
    """
    return html


def format_place_html(p, is_future=False):
    name = p.get("displayName", {}).get("text", "Névtelen")
    address = p.get("formattedAddress", "")
    maps_url = p.get("googleMapsUri", "")
    website = p.get("websiteUri", "")
    phone = p.get("nationalPhoneNumber", "")
    place_type = p.get("primaryType", "")
    area = p.get("_search_area", "")

    badge = ""
    extra = ""
    if is_future:
        badge = '<span class="badge future">HAMAROSAN NYÍL</span>'
        opening = p.get("openingDate", {})
        if opening:
            y = opening.get("year", "")
            m = opening.get("month", "")
            d = opening.get("day", "")
            extra = f"<div class='meta'>📅 Várható nyitás: {y}-{m:02d}-{d:02d}</div>" if m else ""
    else:
        badge = '<span class="badge new">ÚJ</span>'

    css_class = "place future" if is_future else "place"

    html = f'<div class="{css_class}">'
    html += f'<div class="name">{badge} {name}</div>'
    if address:
        html += f'<div class="meta">📍 {address}</div>'
    if place_type:
        html += f'<div class="meta">🏷️ {place_type} • {area}</div>'
    if phone:
        html += f'<div class="meta">📞 {phone}</div>'
    if website:
        html += f'<div class="meta">🌐 <a href="{website}">{website}</a></div>'
    if extra:
        html += extra
    if maps_url:
        html += f'<div class="meta"><a href="{maps_url}">▶ Megnyit Google Maps-ben</a></div>'
    html += "</div>"
    return html


def send_email(html_content):
    """E-mail küldés Gmail SMTP-n keresztül."""
    if not (EMAIL_FROM and EMAIL_TO and EMAIL_PASSWORD):
        print("⚠️  E-mail credentials hiányoznak, kihagyom a küldést.")
        print("=" * 60)
        print("HTML ELŐNÉZET (első 2000 karakter):")
        print(html_content[:2000])
        return

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"🍕 Foodora Lead Finder – {datetime.now().strftime('%Y-%m-%d')}"
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg.attach(MIMEText(html_content, "html", "utf-8"))

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
        server.login(EMAIL_FROM, EMAIL_PASSWORD)
        server.send_message(msg)
    print(f"✉️  E-mail elküldve: {EMAIL_TO}")


# ============================================================
# ADATTÁROLÁS - mit láttunk már
# ============================================================

def load_known(path):
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_known(path, items):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(sorted(list(items)), f, ensure_ascii=False, indent=2)


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 60)
    print(f"🚀 Foodora Lead Finder – {datetime.now().isoformat()}")
    print("=" * 60)

    # 1. Google Places
    print("\n📍 Google Places API lekérdezés...")
    known_place_ids = load_known(KNOWN_PLACES_FILE)
    all_places = fetch_google_places()
    new_places = filter_new_places(all_places, known_place_ids)
    print(f"  → {len(new_places)} ÚJ hely (összesen ismert: {len(known_place_ids)})")

    # 2. Új domainek
    print("\n🌐 crt.sh új domainek lekérdezése...")
    known_domains = load_known(KNOWN_DOMAINS_FILE)
    found_domains = fetch_new_domains()
    new_domains = [d for d in found_domains if d not in known_domains]
    print(f"  → {len(new_domains)} ÚJ domain")

    # 3. Hírek
    print("\n📰 Google News RSS...")
    news = fetch_news()
    print(f"  → {len(news)} friss hír")

    # 4. We Love Budapest gasztro cikkek
    print("\n🌶️  We Love Budapest gasztro cikkek...")
    known_wlb = load_known(KNOWN_WLB_FILE)
    wlb_articles = fetch_welovebudapest()
    new_wlb_articles = filter_new_articles(wlb_articles, known_wlb)
    print(f"  → {len(new_wlb_articles)} ÚJ cikk (összesen ismert: {len(known_wlb)})")

    # 5. Funzine gasztro cikkek
    print("\n🎉 Funzine gasztro cikkek...")
    known_funzine = load_known(KNOWN_FUNZINE_FILE)
    funzine_articles = fetch_funzine()
    new_funzine_articles = filter_new_articles(funzine_articles, known_funzine)
    print(f"  → {len(new_funzine_articles)} ÚJ cikk (összesen ismert: {len(known_funzine)})")

    # 6. Time Out Budapest cikkek
    print("\n🌍 Time Out Budapest cikkek...")
    known_timeout = load_known(KNOWN_TIMEOUT_FILE)
    timeout_articles = fetch_timeout_budapest()
    new_timeout_articles = filter_new_articles(timeout_articles, known_timeout)
    print(f"  → {len(new_timeout_articles)} ÚJ cikk (összesen ismert: {len(known_timeout)})")

    # 7. E-mail
    print("\n✉️  E-mail összeállítása és küldése...")
    html = format_email(
        new_places, new_domains, news,
        new_wlb_articles, new_funzine_articles, new_timeout_articles,
    )
    send_email(html)

    # 8. Mentés
    known_place_ids.update(p.get("id") for p in all_places if p.get("id"))
    known_domains.update(found_domains)
    known_wlb.update(a["url"] for a in wlb_articles)
    known_funzine.update(a["url"] for a in funzine_articles)
    known_timeout.update(a["url"] for a in timeout_articles)
    save_known(KNOWN_PLACES_FILE, known_place_ids)
    save_known(KNOWN_DOMAINS_FILE, known_domains)
    save_known(KNOWN_WLB_FILE, known_wlb)
    save_known(KNOWN_FUNZINE_FILE, known_funzine)
    save_known(KNOWN_TIMEOUT_FILE, known_timeout)
    print(
        f"\n💾 Adatok elmentve: {len(known_place_ids)} hely, "
        f"{len(known_domains)} domain, "
        f"{len(known_wlb)} WLB, {len(known_funzine)} Funzine, "
        f"{len(known_timeout)} Time Out cikk"
    )
    print("✅ Kész!")


if __name__ == "__main__":
    main()
