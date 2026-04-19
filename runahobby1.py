import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import (
    Flask,
    flash,
    make_response,
    redirect,
    render_template,
    request,
    url_for,
)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "Hobbys.xlsx"
DB_PATH = BASE_DIR / "rentahobby.db"

LANG_TEXTS = {
    "de": {
        "title": "Rent a Hobby",
        "home": "Home",
        "search": "Search",
        "hobbies": "Hobbies",
        "faqs": "FAQs",
        "contact": "Contact",
        "hero_title": "RENT A HOBBY",
        "hero_subtitle": "Probieren Sie etwas Neues aus - ohne Verpflichtung!",
        "search_placeholder": "Suche nach einem Hobby...",
        "find_hobby": "Finde ein Hobby",
        "popular_hobbies": "Beliebte Hobbys",
        "price": "Preis",
        "more_info": "Mehr erfahren",
        "no_match": "Wollen Sie, dass wir dieses Hobby in Zukunft anbieten?",
        "yes": "Ja",
        "no": "Nein",
        "accessories": "Zubehör",
        "detail_text": "Wir arbeiten daran, unsere Hobbys bald zur Verfügung zu stellen. Bitte geben Sie Ihre E-Mail-Adresse ein, und wir melden uns bei Ihnen, sobald es losgeht.",
        "buy_button": "Kaufen",
        "email_label": "Ihre E-Mail-Adresse",
        "note_title": "Hinweis",
        "note_text": "Lieber Kunde, unser Startup ist noch in der Validierungsphase. Wir sind noch nicht bereit, Vermietungen anzubieten, aber wir melden uns bei Ihnen, sobald wir starten.",
        "back": "Zurück zu Hobbys",
        "contact_title": "Kontakt",
        "contact_support": "Kontaktieren Sie uns gerne per E-Mail oder Telefon.",
        "contact_text": "Wir freuen uns auf Ihre Nachricht und melden uns schnellstmöglich zurück.",
        "phone_label": "Telefon",
        "faq_question_1": "Wie funktioniert Rent a Hobby?",
        "faq_answer_1": "Wir listen Hobbys, geben Beschreibungen und Preise an und nehmen Interessenten auf, damit wir unser Angebot verbessern können.",
        "faq_question_2": "Kann ich sofort reservieren?",
        "faq_answer_2": "Noch nicht. Unsere Plattform ist in der Validierungsphase. Tragen Sie Ihre E-Mail ein, und wir informieren Sie, sobald Vermietungen möglich sind.",
        "faq_question_3": "Wie kann ich ein neues Hobby vorschlagen?",
        "faq_answer_3": "Nutzen Sie die Suche und geben Sie das Hobby ein. Falls es noch nicht vorhanden ist, können Sie uns mitteilen, dass Sie es gerne hätten.",
        "title_home": "Startseite",
    },
    "en": {
        "title": "Rent a Hobby",
        "home": "Home",
        "search": "Search",
        "hobbies": "Hobbies",
        "faqs": "FAQs",
        "contact": "Contact",
        "hero_title": "RENT A HOBBY",
        "hero_subtitle": "Try something new — no commitment!",
        "search_placeholder": "Search for a hobby...",
        "find_hobby": "Find a Hobby",
        "popular_hobbies": "Popular Hobbies",
        "price": "Price",
        "more_info": "Learn more",
        "no_match": "Would you like us to offer this hobby in the future?",
        "yes": "Yes",
        "no": "No",
        "accessories": "Accessories",
        "detail_text": "We are still validating our startup. Enter your email and we will contact you as soon as we are ready.",
        "buy_button": "Buy",
        "email_label": "Your email address",
        "note_title": "Note",
        "note_text": "Dear customer, our startup is not ready to sell or rent yet. We are still validating the idea and will contact you once we are ready.",
        "back": "Back to Hobbies",
        "contact_title": "Contact",
        "contact_support": "Please contact us by email or phone.",
        "contact_text": "We will get back to you as soon as possible.",
        "phone_label": "Phone",
        "faq_question_1": "How does Rent a Hobby work?",
        "faq_answer_1": "We list hobbies with descriptions and prices, and collect interested users to improve our service.",
        "faq_question_2": "Can I reserve immediately?",
        "faq_answer_2": "Not yet. Our platform is in the validation phase. Please leave your email and we will notify you when rentals are available.",
        "faq_question_3": "How can I suggest a new hobby?",
        "faq_answer_3": "Use the search field and type the hobby. If it is not available, you can tell us and we will consider adding it.",
        "title_home": "Home",
    },
}

app = Flask(__name__)
app.secret_key = "change-this-secret"


def get_lang():
    lang = request.cookies.get("lang", "de")
    return lang if lang in LANG_TEXTS else "de"


def load_hobbies():
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.execute("SELECT * FROM hobbies")
        rows = c.fetchall()
        return [{"Hobby": r[0], "Beschreibung Hobby": r[1], "Zubehör": r[2], "Preis": r[3], "image": r[4]} for r in rows]


def get_hobby_image(hobby_name):
    # Mapping for hobby names to image filenames
    image_mapping = {
        "Fotographie": "Fotografie.avif",
        "Keramik": "Keramik.avif",
        "Bowling": "Bowling.avif",
        "saxophone": "saxofon.avif",
        "E-Guitarre": "eguittare.avif",
        "3D-Druck": "3d-druck.avif",
        "Bogenschiessen": "Bogenschiessen.avif",
        "Drohnen fliegen": "drohnen fliegen.avif",
        "Escape Room": "escape room.avif",
        "Modelbauauto": "modelbauauto.avif",
        "Schmuckherstellung": "Schmuckherstellung.avif",
        "Siebdruck": "siebdruck.avif",
        "Standup Paddling": "standup paddling.avif",
        "Teleskop": "teleskop.avif",
        "Virtual Reality": "virtual reality.avif",
    }
    return image_mapping.get(hobby_name, "placeholder.png")  # fallback if no image found


def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS hobbies (Hobby TEXT PRIMARY KEY, Beschreibung TEXT, Zubehör TEXT, Preis INTEGER, image TEXT)")
        # Lade Daten aus Excel, falls DB leer
        c.execute("SELECT COUNT(*) FROM hobbies")
        if c.fetchone()[0] == 0:
            hobbies = load_hobbies()
            for hobby in hobbies:
                c.execute("INSERT INTO hobbies VALUES (?, ?, ?, ?, ?)", 
                          (hobby["Hobby"], hobby["Beschreibung Hobby"], hobby["Zubehör"], hobby["Preis"], hobby["image"]))
        # Bestehende Tabellen
        c.execute(
            "CREATE TABLE IF NOT EXISTS wishlist (id INTEGER PRIMARY KEY, hobby TEXT, created_at TEXT)"
        )
        c.execute(
            "CREATE TABLE IF NOT EXISTS interests (id INTEGER PRIMARY KEY, hobby TEXT, email TEXT, created_at TEXT)"
        )
        c.execute(
            "CREATE TABLE IF NOT EXISTS hobby_views (hobby TEXT PRIMARY KEY, view_count INTEGER DEFAULT 0)"
        )
        conn.commit()


def find_hobby(name):
    if not name:
        return None
    normalized = name.strip().lower()
    for hobby in load_hobbies():
        if hobby["Hobby"].strip().lower() == normalized:
            return hobby
    return None


def search_hobbies(query):
    if not query:
        return []
    needle = query.strip().lower()
    results = []
    for hobby in load_hobbies():
        text = " ".join(
            [hobby["Hobby"], hobby["Beschreibung Hobby"], hobby["Zubehör"]]
        ).lower()
        if needle in text:
            results.append(hobby)
    return results


def store_request(hobby_name):
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.execute(
            "INSERT INTO wishlist (hobby, created_at) VALUES (?, ?)",
            (hobby_name.strip(), datetime.utcnow().isoformat()),
        )
        conn.commit()


def store_interest(hobby_name, email):
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.execute(
            "INSERT INTO interests (hobby, email, created_at) VALUES (?, ?, ?)",
            (hobby_name.strip(), email.strip(), datetime.utcnow().isoformat()),
        )
        conn.commit()


def increment_hobby_view(hobby_name):
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.execute(
            "INSERT OR IGNORE INTO hobby_views (hobby, view_count) VALUES (?, 0)",
            (hobby_name.strip(),),
        )
        c.execute(
            "UPDATE hobby_views SET view_count = view_count + 1 WHERE hobby = ?",
            (hobby_name.strip(),),
        )
        conn.commit()


def get_hobby_views(hobby_name):
    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        c.execute("SELECT view_count FROM hobby_views WHERE hobby = ?", (hobby_name.strip(),))
        result = c.fetchone()
        return result[0] if result else 0


def page_context(title_key="title"):
    lang = get_lang()
    texts = LANG_TEXTS[lang]
    texts = {**texts, "title": texts.get(title_key, texts["title"])}
    return lang, texts


@app.route("/")
def home():
    lang, texts = page_context("title_home")
    hobbies = load_hobbies()
    featured = hobbies[:4]
    return render_template("home.html", texts=texts, featured=featured, lang=lang)


@app.route("/search")
def search():
    lang, texts = page_context("search")
    query = request.args.get("q")
    results = search_hobbies(query) if query else []
    return render_template(
        "search.html",
        texts=texts,
        query=query,
        results=results,
        lang=lang,
    )


@app.route("/hobbies")
@app.route("/find")
def hobbies():
    lang, texts = page_context("hobbies")
    return render_template("hobbies.html", texts=texts, hobbies=load_hobbies(), lang=lang)


@app.route("/hobby/<path:hobby_name>")
def hobby_detail(hobby_name):
    lang, texts = page_context("hobbies")
    hobby = find_hobby(hobby_name)
    if hobby is None:
        flash(texts["no_match"], "error")
        return redirect(url_for("hobbies"))
    increment_hobby_view(hobby_name)
    return render_template("detail.html", texts=texts, hobby=hobby, lang=lang)


@app.route("/request_hobby", methods=["POST"])
def request_hobby():
    lang = get_lang()
    texts = LANG_TEXTS[lang]
    hobby_name = request.form.get("hobby_name", "")
    answer = request.form.get("answer")
    if answer == "yes" and hobby_name.strip():
        store_request(hobby_name)
        flash(texts["buy_button"] + ": " + hobby_name + " - " + texts["note_text"], "success")
    else:
        flash(texts["no_match"], "error")
    return redirect(url_for("search", q=hobby_name))


@app.route("/express_interest", methods=["POST"])
def express_interest():
    lang = get_lang()
    texts = LANG_TEXTS[lang]
    hobby_name = request.form.get("hobby_name", "")
    email = request.form.get("email", "")
    if email.strip() and hobby_name.strip():
        store_interest(hobby_name, email)
        flash(
            "{}: {}. {}".format(
                texts["buy_button"], hobby_name, texts["note_text"]
            ),
            "success",
        )
        return redirect(url_for("hobby_detail", hobby_name=hobby_name))
    flash(texts["email_label"] + " fehlt", "error")
    return redirect(url_for("hobby_detail", hobby_name=hobby_name))


@app.route("/faqs")
def faqs():
    lang, texts = page_context("faqs")
    return render_template("faqs.html", texts=texts, lang=lang)


@app.route("/contact")
def contact():
    lang, texts = page_context("contact")
    return render_template("contact.html", texts=texts, lang=lang)


@app.route("/set_language/<lang>")
def set_language(lang):
    if lang not in LANG_TEXTS:
        lang = "de"
    next_page = request.referrer or url_for("home")
    response = make_response(redirect(next_page))
    response.set_cookie("lang", lang, max_age=60 * 60 * 24 * 365)
    return response


if __name__ == "__main__":
    init_db()
    app.run()
