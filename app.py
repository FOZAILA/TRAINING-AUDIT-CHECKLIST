"""HSE Compliance Checklist website.
Upload the Master Training Register -> the HSE Compliance Checklist is filled automatically.
Run:  python app.py   then open http://127.0.0.1:5000"""
import datetime as dt
import io
import json
import os

from flask import Flask, Response, flash, redirect, render_template, request, send_file, url_for

from core.checklist import fill_checklist
from core.config import BASE_DIR, DATA_DIR, FORMATS_DIR, load_settings
from core.register import parse_register, resolve_as_of

app = Flask(__name__)
app.secret_key = "hse-checklist-local"
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

REG_PATH = os.path.join(DATA_DIR, "latest_register.xlsx")
META_PATH = os.path.join(DATA_DIR, "latest_register.json")
_cache = {}


@app.before_request
def _password_gate():
    """If the environment variable APP_PASSWORD is set (hosted use), ask for it (user name: anything)."""
    pw = os.environ.get("APP_PASSWORD")
    if not pw:
        return None
    auth = request.authorization
    if auth and auth.password == pw:
        return None
    return Response("Password required", 401, {"WWW-Authenticate": 'Basic realm="HSE Checklist"'})


def _meta():
    if os.path.exists(META_PATH):
        with open(META_PATH, encoding="utf-8") as fh:
            return json.load(fh)
    return None


def current():
    """Parse the LAST uploaded register (cached until a new file is uploaded)."""
    meta = _meta()
    if not meta or not os.path.exists(REG_PATH):
        return None, None
    settings = load_settings()
    stamp = (meta["uploaded_at"], settings["as_of_date"], os.path.getmtime(os.path.join(BASE_DIR, "config", "settings.json")))
    if _cache.get("stamp") != stamp:
        reg = parse_register(REG_PATH, settings, meta["filename"])
        _cache.update(stamp=stamp, reg=reg, as_of=resolve_as_of(settings))
    return _cache["reg"], _cache["as_of"]


def _css():
    with open(os.path.join(BASE_DIR, "static", "style.css"), encoding="utf-8") as fh:
        return fh.read()


@app.route("/")
def index():
    reg, _ = current()
    return render_template("index.html", meta=_meta(), reg=reg, css=_css())


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("register")
    if not f or not f.filename.lower().endswith((".xlsx", ".xlsm")):
        flash("Please choose the Master Training Register (.xlsx).")
        return redirect(url_for("index"))
    os.makedirs(DATA_DIR, exist_ok=True)
    f.save(REG_PATH)
    try:
        parse_register(REG_PATH, load_settings(), f.filename)
    except Exception as exc:                          # bad file: keep nothing half-loaded
        os.remove(REG_PATH)
        flash("Could not read this register: %s" % exc)
        return redirect(url_for("index"))
    with open(META_PATH, "w", encoding="utf-8") as fh:
        json.dump({"filename": f.filename, "uploaded_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}, fh)
    _cache.clear()
    flash("Register loaded: %s" % f.filename)
    return redirect(url_for("index"))


@app.route("/checklist")
def checklist_page():
    reg, as_of = current()
    if reg is None:
        return redirect(url_for("index"))
    s = load_settings()
    _, info = fill_checklist(reg, s, as_of, os.path.join(FORMATS_DIR, s["checklist"]["template"]))
    return render_template("checklist.html", info=info, as_of=as_of, meta=_meta(), css=_css())


@app.route("/download/checklist.xlsx")
def dl_checklist():
    reg, as_of = current()
    if reg is None:
        return redirect(url_for("index"))
    s = load_settings()
    data, _ = fill_checklist(reg, s, as_of, os.path.join(FORMATS_DIR, s["checklist"]["template"]),
                             audit_no=request.args.get("audit_no") or None)
    return send_file(io.BytesIO(data), as_attachment=True,
                     download_name="HSE_Compliance_Checklist_%s.xlsx" % as_of.strftime("%d-%m-%Y"),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    # local use: 127.0.0.1.  Office network: set HOST=0.0.0.0.  Hosting services set PORT themselves.
    app.run(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", 5000)), debug=False)
