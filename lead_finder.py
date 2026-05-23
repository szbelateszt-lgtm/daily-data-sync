"""
Foodora Lead Finder - Új éttermek és élelmiszerboltok automatikus felderítése
Budapest + agglomeráció területén.

Adatforrások:
1. Google Places API (New) - hivatalos, megbízható
2. crt.sh - új .hu domainek vendéglátós kulcsszavakkal
3. Google News RSS - friss hírek új helyekről

A script naponta egyszer fut le GitHub Actions-ben.
Az új találatokat e-mailben küldi el.
"""

import os
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
        for query in ["étterem", "kávézó", "élelmiszerbolt", "pékség", "bisztró"]:
            body = {
                "textQuery": f"{query} {area['name']}",
                "includedType": "restaurant",  # tág értelmezésű
                "includePureServiceAreaBusinesses": False,
                "locationBias": {
                    "circle": {
                        "center": {
                            "latitude": area["lat"],
                            "longitude": area["lng"],
                        },
                        "radius": area["radius"],
                    }
                },
                "pageSize": 20,
                "includedFutureOpening": True,  # KULCSFONTOSSÁGÚ!
            }

            try:
                resp = requests.post(
                    "https://places.googleapis.com/v1/places:searchText",
                    headers=headers,
                    json=body,
                    timeout=30,
                )
                resp.raise_for_status()
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
    cutoff = datetime.utcnow() - timedelta(days=2)

    for q in queries:
        url = (
            f"https://news.google.com/rss/search?q={requests.utils.quote(q)}"
            f"&hl=hu&gl=HU&ceid=HU:hu"
        )
        try:
            feed = feedparser.parse(url)
            for entry in feed.entries[:10]:
                pub = entry.get("published_parsed")
                if pub:
                    pub_dt = datetime(*pub[:6])
                    if pub_dt < cutoff:
                        continue
                all_news.append({
                    "title": entry.get("title", ""),
                    "link": entry.get("link", ""),
                    "published": entry.get("published", ""),
                    "source": entry.get("source", {}).get("title", "") if entry.get("source") else "",
                    "query": q,
                })
            print(f"  ✓ '{q}': {len(feed.entries[:10])} hír")
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
# E-MAIL FORMÁZÁS ÉS KÜLDÉS
# ============================================================

def format_email(new_places, new_domains, news):
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
      Hírek: {len(news)}
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

    html += """
    <hr>
    <p style="color: #999; font-size: 11px;">
      Ezt az e-mailt a Foodora Lead Finder automata küldte.
      Adatforrások: Google Places API, crt.sh, Google News.
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

    # 4. E-mail
    print("\n✉️  E-mail összeállítása és küldése...")
    html = format_email(new_places, new_domains, news)
    send_email(html)

    # 5. Mentés
    known_place_ids.update(p.get("id") for p in all_places if p.get("id"))
    known_domains.update(found_domains)
    save_known(KNOWN_PLACES_FILE, known_place_ids)
    save_known(KNOWN_DOMAINS_FILE, known_domains)
    print(f"\n💾 Adatok elmentve: {len(known_place_ids)} hely, {len(known_domains)} domain")
    print("✅ Kész!")


if __name__ == "__main__":
    main()
