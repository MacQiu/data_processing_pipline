import json
import os
import sys
import signal
import threading
import time
import webbrowser

from flask import Flask, jsonify, render_template, request

from extract import extract_all, run as run_extraction

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------


def _get_template_dir():
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, "templates")
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")


HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "..", "data")
JSON_PATH = os.path.join(HERE, "performance_data.json")

app = Flask(__name__, template_folder=_get_template_dir())

# In-memory cache loaded from JSON
_data = {}


def _load_data():
    global _data
    if not os.path.exists(JSON_PATH):
        run_extraction(DATA_DIR, JSON_PATH)
    with open(JSON_PATH, "r", encoding="utf-8") as f:
        _data = json.load(f)


_load_data()

# ---------------------------------------------------------------------------
# Heartbeat / shutdown (same pattern as the main app)
# ---------------------------------------------------------------------------
_last_heartbeat = time.time()
_HEARTBEAT_TIMEOUT = 120
_shutdown_requested = False


def _watchdog():
    while True:
        time.sleep(10)
        if _shutdown_requested:
            os.kill(os.getpid(), signal.SIGTERM)
            return
        if time.time() - _last_heartbeat > _HEARTBEAT_TIMEOUT:
            os.kill(os.getpid(), signal.SIGTERM)
            return


@app.route("/api/heartbeat", methods=["POST"])
def heartbeat():
    global _last_heartbeat
    _last_heartbeat = time.time()
    return "", 204


@app.route("/api/shutdown", methods=["POST"])
def shutdown():
    global _shutdown_requested
    _shutdown_requested = True
    return "", 204


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _ym_key(year, month):
    return year * 100 + month


def _parse_ym(s):
    """Parse 'YYYY-MM' into (year, month) ints."""
    parts = s.split("-")
    return int(parts[0]), int(parts[1])


def _in_range(year, month, start, end):
    k = _ym_key(year, month)
    return _ym_key(*start) <= k <= _ym_key(*end)


# ---------------------------------------------------------------------------
# API endpoints
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/time-range")
def api_time_range():
    """Return the min/max year-month present in the data."""
    all_ym = set()
    for recs in _data.get("people", {}).values():
        for r in recs:
            all_ym.add((r["year"], r["month"]))
    for recs in _data.get("office_scores", {}).values():
        for r in recs:
            all_ym.add((r["year"], r["month"]))

    if not all_ym:
        return jsonify({"months": []})

    sorted_ym = sorted(all_ym)
    return jsonify({
        "months": [{"year": y, "month": m} for y, m in sorted_ym],
    })


@app.route("/api/offices")
def api_offices():
    """Return offices that have data in the given time range."""
    start_s = request.args.get("start")
    end_s = request.args.get("end")

    offices = set()

    if start_s and end_s:
        start, end = _parse_ym(start_s), _parse_ym(end_s)
        for name, recs in _data.get("people", {}).items():
            for r in recs:
                if _in_range(r["year"], r["month"], start, end):
                    offices.add(r["office"])
        for office, recs in _data.get("office_scores", {}).items():
            for r in recs:
                if _in_range(r["year"], r["month"], start, end):
                    offices.add(office)
    else:
        for recs in _data.get("people", {}).values():
            for r in recs:
                offices.add(r["office"])
        for office in _data.get("office_scores", {}):
            offices.add(office)

    return jsonify({"offices": sorted(offices)})


@app.route("/api/people")
def api_people():
    """Return people who worked at a given office in the time range."""
    office = request.args.get("office", "")
    start_s = request.args.get("start")
    end_s = request.args.get("end")

    results = []

    for name, recs in _data.get("people", {}).items():
        matched = False
        ptype = "employee"
        for r in recs:
            in_time = True
            if start_s and end_s:
                start, end = _parse_ym(start_s), _parse_ym(end_s)
                in_time = _in_range(r["year"], r["month"], start, end)
            if in_time and r["office"] == office:
                matched = True
                ptype = r["type"]
        if matched:
            results.append({"name": name, "type": ptype})

    results.sort(key=lambda x: (0 if x["type"] == "manager" else 1, x["name"]))
    return jsonify({"people": results})


@app.route("/api/person/<name>")
def api_person(name):
    """Return a person's monthly data for chart rendering."""
    start_s = request.args.get("start")
    end_s = request.args.get("end")

    recs = _data.get("people", {}).get(name, [])
    if not recs:
        return jsonify({"error": "Person not found"}), 404

    filtered = []
    for r in recs:
        if start_s and end_s:
            start, end = _parse_ym(start_s), _parse_ym(end_s)
            if not _in_range(r["year"], r["month"], start, end):
                continue
        filtered.append(r)

    filtered.sort(key=lambda r: _ym_key(r["year"], r["month"]))

    labels = []
    marks = []
    result_counts = {"A": 0, "B": 0, "C": 0}
    details = []

    for r in filtered:
        label = f"{r['year']}-{r['month']:02d}"
        labels.append(label)
        marks.append(r["mark"])
        if r["result"] in result_counts:
            result_counts[r["result"]] += 1
        details.append({
            "month": label,
            "office": r["office"],
            "type": r["type"],
            "mark": r["mark"],
            "result": r["result"],
        })

    return jsonify({
        "name": name,
        "labels": labels,
        "marks": marks,
        "result_counts": result_counts,
        "details": details,
    })


@app.route("/api/refresh", methods=["POST"])
def api_refresh():
    """Re-extract data from Excel files and reload the cache."""
    try:
        run_extraction(DATA_DIR, JSON_PATH)
        _load_data()
        return jsonify({
            "ok": True,
            "files": _data.get("files_processed", []),
            "people_count": len(_data.get("people", {})),
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _open_browser():
    webbrowser.open("http://127.0.0.1:5001")


if __name__ == "__main__":
    wd = threading.Thread(target=_watchdog, daemon=True)
    wd.start()

    threading.Timer(1.5, _open_browser).start()
    app.run(host="127.0.0.1", port=5001, debug=False)
