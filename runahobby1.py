import os
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
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "Hobbys.xlsx"

# Database Configuration - Local SQLite or Remote PostgreSQL
DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    f"sqlite:///{BASE_DIR}/rentahobby.db"
)

# Fix PostgreSQL URL if needed
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.secret_key = "change-this-secret"

db = SQLAlchemy(app)


# Database Models
class Hobby(db.Model):
    __tablename__ = "hobbies"
    id = db.Column(db.Integer, primary_key=True)
    hobby = db.Column(db.String(255), unique=True, nullable=False)
    beschreibung = db.Column(db.Text)
    zubehor = db.Column(db.Text)
    preis = db.Column(db.Integer)
    image = db.Column(db.String(255))

    def to_dict(self):
        return {
            "Hobby": self.hobby,
            "Beschreibung Hobby": self.beschreibung or "",
            "Zubehör": self.zubehor or "",
            "Preis": self.preis or 0,
            "image": self.image or "",
        }


class Interest(db.Model):
    __tablename__ = "interests"
    id = db.Column(db.Integer, primary_key=True)
    hobby = db.Column(db.String(255))
    email = db.Column(db.String(255))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class HobbyView(db.Model):
    __tablename__ = "hobby_views"
    id = db.Column(db.Integer, primary_key=True)
    hobby = db.Column(db.String(255), unique=True)
    view_count = db.Column(db.Integer, default=0)


class Wishlist(db.Model):
    __tablename__ = "wishlist"
    id = db.Column(db.Integer, primary_key=True)
    hobby = db.Column(db.String(255))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# Language texts
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


def get_lang():
    lang = request.cookies.get("lang", "de")
    return lang if lang in LANG_TEXTS else "de"


def normalize_name(value):
    if not value:
        return ""
    normalized = str(value).lower().strip()
    normalized = normalized.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    normalized = normalized.replace("é", "e").replace("è", "e").replace("ê", "e")
    normalized = normalized.replace(" ", "").replace("-", "").replace("_", "").replace(".", "")
    return "".join(ch for ch in normalized if ch.isalnum())


def get_hobby_image(hobby_name):
    if not hobby_name:
        return ""

    static_dir = BASE_DIR / "static"
    try:
        available_files = [
            path.name
            for path in static_dir.iterdir()
            if path.is_file() and path.suffix.lower() in {".avif", ".png", ".jpg", ".jpeg", ".webp"}
        ]
    except (FileNotFoundError, OSError):
        available_files = []

    normalized_hobby = normalize_name(hobby_name)
    for filename in available_files:
        if normalize_name(Path(filename).stem) == normalized_hobby:
            return filename

    image_mapping = {
        "fotografie": "Fotografie.avif",
        "fotographie": "Fotografie.avif",
        "keramik": "Keramik.avif",
        "bowling": "Bowling.avif",
        "saxophone": "saxofon.avif",
        "saxophon": "saxofon.avif",
        "egitarre": "eguittare.avif",
        "eguitarre": "eguittare.avif",
        "e-gitarre": "eguittare.avif",
        "3ddruck": "3d-druck.avif",
        "teleskopastronomie": "teleskop.avif",
        "teleskop": "teleskop.avif",
        "siebdruckdiytextildruck": "siebdruck.avif",
        "modelbauauto": "modelbauauto.avif",
        "drohnenfliegen": "drohnen fliegen.avif",
        "virtualreality": "virtual reality.avif",
        "bogenschiessen": "Bogenschiessen.avif",
        "schmuckherstellung": "Schmuckherstellung.avif",
        "standuppaddlingsup": "standup paddling.avif",
        "escaperoomspielefuerzuhause": "escape room.avif",
    }

    mapped = image_mapping.get(normalized_hobby)
    if mapped and mapped in available_files:
        return mapped

    return "placeholder.png" if "placeholder.png" in available_files else (available_files[0] if available_files else "")


def load_hobbies_from_excel():
    if not EXCEL_PATH.exists():
        return []

    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    hobbies = []
    for _, row in df.iterrows():
        hobby_name = str(row.get("Hobby", "")).strip()
        if not hobby_name:
            continue
        hobbies.append(
            {
                "Hobby": hobby_name,
                "Beschreibung Hobby": str(row.get("Beschreibung Hobby", "")).strip(),
                "Zubehör": str(row.get("Zubehör", "")).strip(),
                "Preis": int(row.get("Preis", 0)) if pd.notna(row.get("Preis", 0)) else 0,
                "image": get_hobby_image(hobby_name),
            }
        )
    return hobbies


def init_db():
    with app.app_context():
        db.create_all()
        
        # Load from Excel if DB is empty
        if Hobby.query.count() == 0:
            hobbies = load_hobbies_from_excel()
            for hobby in hobbies:
                h = Hobby(
                    hobby=hobby["Hobby"],
                    beschreibung=hobby["Beschreibung Hobby"],
                    zubehor=hobby["Zubehör"],
                    preis=hobby["Preis"],
                    image=hobby["image"],
                )
                db.session.add(h)
            db.session.commit()
        
        # Fix wrong images
        for hobby in Hobby.query.all():
            correct_image = get_hobby_image(hobby.hobby)
            if correct_image and hobby.image != correct_image:
                hobby.image = correct_image
                db.session.commit()


def load_hobbies():
    hobbies = Hobby.query.all()
    return [h.to_dict() for h in hobbies]


def find_hobby(name):
    if not name:
        return None
    hobby = Hobby.query.filter(func.lower(Hobby.hobby) == name.strip().lower()).first()
    return hobby.to_dict() if hobby else None


def search_hobbies(query):
    if not query:
        return []
    needle = query.strip().lower()
    hobbies = Hobby.query.all()
    results = []
    for hobby in hobbies:
        text = f"{hobby.hobby} {hobby.beschreibung} {hobby.zubehor}".lower()
        if needle in text:
            results.append(hobby.to_dict())
    return results


def store_request(hobby_name):
    w = Wishlist(hobby=hobby_name.strip())
    db.session.add(w)
    db.session.commit()


def store_interest(hobby_name, email):
    i = Interest(hobby=hobby_name.strip(), email=email.strip())
    db.session.add(i)
    db.session.commit()


def increment_hobby_view(hobby_name):
    view = HobbyView.query.filter_by(hobby=hobby_name.strip()).first()
    if view:
        view.view_count += 1
    else:
        view = HobbyView(hobby=hobby_name.strip(), view_count=1)
        db.session.add(view)
    db.session.commit()


def get_hobby_views(hobby_name):
    view = HobbyView.query.filter_by(hobby=hobby_name.strip()).first()
    return view.view_count if view else 0


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


@app.route("/admin")
def admin():
    lang, texts = page_context("contact")
    interests = Interest.query.order_by(Interest.created_at.desc()).all()
    views = HobbyView.query.order_by(HobbyView.view_count.desc()).all()
    wishlist = Wishlist.query.order_by(Wishlist.created_at.desc()).all()
    
    interests_data = [(i.id, i.hobby, i.email, i.created_at) for i in interests]
    views_data = [(v.hobby, v.view_count) for v in views]
    wishlist_data = [(w.id, w.hobby, w.created_at) for w in wishlist]
    
    return render_template("admin.html", texts=texts, interests=interests_data, views=views_data, wishlist=wishlist_data, lang=lang)


if __name__ == "__main__":
    init_db()
    app.run()
