import hmac
import secrets
from datetime import date
from functools import wraps

from flask import (Blueprint, abort, current_app, flash, jsonify, redirect,
                   render_template, request, send_file, session, url_for)

from . import calc, db, excel
from .fields import (BY_KEY, COMPACT_KEYS, FIELDS, FORM_SECTIONS,
                     INPUT_FIELDS, INVOICE_STATUSES, OPEN_WO_STATUSES,
                     PAYMENT_METHODS, SUGGEST_KEYS, WO_STATUSES)
from .values import InvalidValue, display, form_value, parse

bp = Blueprint("main", __name__)

FLAGS = {
    "open": "Open work orders",
    "needs_input": "Needs financial input",
    "owed": "Balance outstanding",
    "overdue": "Overdue invoices",
    "nte": "Over NTE",
}


# ------------------------------------------------------------ security

def register_template_helpers(app):
    @app.context_processor
    def helpers():
        return {
            "csrf_token": _csrf_token,
            "display": display,
            "form_value": form_value,
            "F": BY_KEY,
            "login_enabled": bool(current_app.config["APP_PASSWORD"]),
        }


def _csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)
    return session["csrf"]


@bp.before_app_request
def guard():
    pw = current_app.config["APP_PASSWORD"]
    if pw and not session.get("signed_in") and request.endpoint not in (
            "main.login", "static"):
        return redirect(url_for("main.login", next=request.full_path))
    if request.method == "POST" and not current_app.config.get("TESTING_NO_CSRF"):
        sent = request.form.get("csrf") or request.headers.get("X-CSRF-Token", "")
        if not hmac.compare_digest(sent, session.get("csrf", "")):
            abort(400, "Your session expired. Go back, refresh the page and try again.")


@bp.route("/login", methods=["GET", "POST"])
def login():
    pw = current_app.config["APP_PASSWORD"]
    if not pw:
        return redirect(url_for("main.dashboard"))
    if request.method == "POST":
        if hmac.compare_digest(request.form.get("password", ""), pw):
            session["signed_in"] = True
            nxt = request.args.get("next") or ""
            return redirect(nxt if nxt.startswith("/") and not nxt.startswith("//")
                            else url_for("main.dashboard"))
        flash("Incorrect password.", "error")
    return render_template("login.html")


@bp.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("main.login"))


# ------------------------------------------------------------ helpers

def _filtered(rows, args):
    q = (args.get("q") or "").strip().lower()
    if q:
        text_keys = [f.key for f in INPUT_FIELDS if f.kind in ("text", "textarea")]
        rows = [r for r in rows
                if any(q in str(r.get(k) or "").lower() for k in text_keys)]
    for key in ("wo_status", "invoice_status", "client", "handler"):
        v = args.get(key)
        if v:
            rows = [r for r in rows if (r.get(key) or "") == v]
    flag = args.get("flag")
    if flag == "open":
        rows = [r for r in rows if r.get("wo_status") in OPEN_WO_STATUSES
                or not r.get("wo_status")]
    elif flag == "needs_input":
        rows = [r for r in rows if str(r["input_status"] or "").startswith("Enter")]
    elif flag == "owed":
        rows = [r for r in rows if r["owed"]]
    elif flag == "overdue":
        rows = [r for r in rows if r["overdue"]]
    elif flag == "nte":
        rows = [r for r in rows if str(r["nte_check"] or "").endswith("exceeds NTE")]
    aging = args.get("aging")
    if aging:
        rows = [r for r in rows if r["aging_bucket"] == aging]
    d_from, d_to = args.get("from"), args.get("to")
    if d_from:
        rows = [r for r in rows if r.get("date_in") and r["date_in"] >= d_from]
    if d_to:
        rows = [r for r in rows if r.get("date_in") and r["date_in"] <= d_to]

    sort = args.get("sort")
    if sort in BY_KEY:
        desc = args.get("dir") == "desc"
        present = [r for r in rows if r.get(sort) not in (None, "", calc.NA)]
        missing = [r for r in rows if r.get(sort) in (None, "", calc.NA)]
        present.sort(key=lambda r: (r[sort].lower() if isinstance(r[sort], str)
                                    else r[sort]), reverse=desc)
        rows = present + missing
    return rows


def _form_data(form):
    data, errors = {}, {}
    for f in INPUT_FIELDS:
        try:
            data[f.key] = parse(f, form.get(f.key), source="form")
        except InvalidValue as e:
            data[f.key] = form.get(f.key)
            errors[f.key] = str(e)
    if not data.get("wo_number"):
        errors["wo_number"] = "WO # is required."
    return data, errors


def _form_context(wo, errors, mode):
    return dict(
        wo=wo, errors=errors, mode=mode, sections=FORM_SECTIONS,
        suggestions={k: db.distinct_values(k) for k in SUGGEST_KEYS},
    )


# ------------------------------------------------------------ pages

@bp.route("/")
def dashboard():
    rows = db.all_work_orders()
    totals = calc.totals(rows)

    by_status = [(s, sum(1 for r in rows if r.get("wo_status") == s)) for s in WO_STATUSES]
    no_status = sum(1 for r in rows if not r.get("wo_status"))
    if no_status:
        by_status.append(("No status", no_status))

    aging = []
    for b in calc.AGING_BUCKETS:
        sel = [r for r in rows if r["aging_bucket"] == b]
        aging.append((b, calc.xround(sum(r["balance_due"] for r in sel)), len(sel)))

    clients = {}
    for r in rows:
        if isinstance(r.get("invoice_revenue"), (int, float)):
            name = r.get("client") or "(no client)"
            clients[name] = clients.get(name, 0) + r["invoice_revenue"]
    top_clients = sorted(clients.items(), key=lambda kv: kv[1], reverse=True)[:8]

    overdue = sorted([r for r in rows if r["overdue"]],
                     key=lambda r: r["days_overdue"], reverse=True)[:8]
    nte = [r for r in rows if str(r["nte_check"] or "").endswith("exceeds NTE")][:8]

    return render_template(
        "dashboard.html", totals=totals, by_status=by_status, aging=aging,
        top_clients=top_clients, overdue=overdue, nte=nte,
        activity=db.activity(limit=10), today=date.today())


@bp.route("/work-orders")
def work_orders():
    all_rows = db.all_work_orders()
    rows = _filtered(all_rows, request.args)
    view = request.args.get("view", "compact")
    columns = FIELDS if view == "all" else [BY_KEY[k] for k in COMPACT_KEYS]
    return render_template(
        "list.html", rows=rows, totals=calc.totals(rows), columns=columns,
        view=view, flags=FLAGS, args=request.args, total_count=len(all_rows),
        wo_statuses=WO_STATUSES, invoice_statuses=INVOICE_STATUSES,
        clients=db.distinct_values("client"), handlers=db.distinct_values("handler"),
        aging_buckets=calc.AGING_BUCKETS)


@bp.route("/work-orders/new", methods=["GET", "POST"])
def new_work_order():
    if request.method == "POST":
        data, errors = _form_data(request.form)
        if not errors and db.find_by_number(data["wo_number"]):
            errors["wo_number"] = f"WO # {data['wo_number']} already exists."
        if not errors:
            wo_id = db.insert(data)
            flash(f"Work order {data['wo_number']} created.", "success")
            return redirect(url_for("main.work_order", wo_id=wo_id))
        flash("Please fix the highlighted fields.", "error")
        return render_template("form.html", **_form_context(data, errors, "new")), 422
    blank = {"wo_status": "New", "invoice_status": "Not invoiced",
             "date_in": date.today().isoformat()}
    copy_id = request.args.get("copy", type=int)
    if copy_id:
        src = db.get_work_order(copy_id)
        if src:
            keep = ["client", "store_name", "location", "address", "job_type",
                    "handler", "tech", "assignees", "target_markup", "payment_terms",
                    "payment_method"]
            blank.update({k: src.get(k) for k in keep})
    return render_template("form.html", **_form_context(blank, {}, "new"))


@bp.route("/work-orders/<int:wo_id>")
def work_order(wo_id):
    wo = db.get_work_order(wo_id) or abort(404)
    return render_template("detail.html", wo=wo, sections=FORM_SECTIONS,
                           activity=db.activity(wo_id), methods=PAYMENT_METHODS,
                           today=date.today().isoformat())


@bp.route("/work-orders/<int:wo_id>/edit", methods=["GET", "POST"])
def edit_work_order(wo_id):
    wo = db.get_work_order(wo_id) or abort(404)
    if request.method == "POST":
        data, errors = _form_data(request.form)
        other = db.find_by_number(data.get("wo_number") or "")
        if not errors and other and other["id"] != wo_id:
            errors["wo_number"] = f"WO # {data['wo_number']} already exists."
        if not errors:
            changed = db.update(wo_id, data)
            flash("Changes saved." if changed else "No changes to save.", "success")
            return redirect(url_for("main.work_order", wo_id=wo_id))
        flash("Please fix the highlighted fields.", "error")
        data["id"] = wo_id
        return render_template("form.html", **_form_context(data, errors, "edit")), 422
    return render_template("form.html", **_form_context(wo, {}, "edit"))


@bp.route("/work-orders/<int:wo_id>/delete", methods=["POST"])
def delete_work_order(wo_id):
    wo = db.get_work_order(wo_id) or abort(404)
    db.delete(wo_id)
    flash(f"Work order {wo['wo_number']} deleted.", "success")
    return redirect(url_for("main.work_orders"))


@bp.route("/work-orders/<int:wo_id>/payment", methods=["POST"])
def record_payment(wo_id):
    """Add a payment to Amount received and set the invoice status to match."""
    wo = db.get_work_order(wo_id) or abort(404)
    try:
        amount = parse(BY_KEY["amount_received"], request.form.get("amount"))
        method = parse(BY_KEY["payment_method"], request.form.get("method"))
    except InvalidValue as e:
        flash(str(e), "error")
        return redirect(url_for("main.work_order", wo_id=wo_id))
    if not amount:
        flash("Enter the amount received.", "error")
        return redirect(url_for("main.work_order", wo_id=wo_id))
    received = calc.xround((wo.get("amount_received") or 0) + amount)
    data = {"amount_received": received,
            "payment_method": method or wo.get("payment_method")}
    revenue = wo.get("invoice_revenue")
    status = wo.get("invoice_status")
    if isinstance(revenue, (int, float)) and status not in ("Void", "Disputed"):
        status = "Paid" if received >= revenue - 0.004 else "Partially paid"
    data["invoice_status"] = status
    db.update(wo_id, data, only=list(data),
              action=f"Payment of ${amount:,.2f} recorded"
                     f" (paid on {request.form.get('paid_on') or date.today().isoformat()})")
    flash(f"Payment of ${amount:,.2f} recorded. Invoice status: {status or '—'}.", "success")
    return redirect(url_for("main.work_order", wo_id=wo_id))


@bp.route("/import", methods=["GET", "POST"])
def import_page():
    if request.method == "GET":
        return render_template("import.html", report=None)
    upload = request.files.get("file")
    if not upload or not upload.filename:
        flash("Choose an Excel (.xlsx) or CSV file to import.", "error")
        return render_template("import.html", report=None), 400
    if not upload.filename.lower().endswith((".xlsx", ".xlsm", ".csv")):
        flash("Only .xlsx, .xlsm or .csv files can be imported.", "error")
        return render_template("import.html", report=None), 400

    report = {"created": [], "updated": [], "unchanged": [], "skipped": [],
              "filename": upload.filename}
    try:
        rows = list(excel.read_rows(upload.filename, upload.stream))
    except Exception as e:      # unreadable or wrong layout
        flash(f"Could not read the file: {e}", "error")
        return render_template("import.html", report=None), 400

    for rownum, raw in rows:
        data, errors = excel.parse_row(raw)
        wo_number = data.get("wo_number")
        if not wo_number:
            report["skipped"].append((rownum, "", ["WO # is empty"]))
            continue
        if errors:
            report["skipped"].append((rownum, wo_number, errors))
            continue
        existing = db.find_by_number(wo_number)
        if existing:
            changed = db.update(existing["id"], data, action="Updated by import",
                                only=list(data))
            report["updated" if changed else "unchanged"].append((rownum, wo_number))
        else:
            db.insert({**{f.key: None for f in INPUT_FIELDS}, **data},
                      action="Imported")
            report["created"].append((rownum, wo_number))
    return render_template("import.html", report=report)


@bp.route("/export.<fmt>")
def export(fmt):
    rows = _filtered(db.all_work_orders(), request.args)
    stamp = date.today().strftime("%Y-%m-%d")
    if fmt == "xlsx":
        rows = sorted(rows, key=lambda r: (r.get("date_in") or "9999", r["id"]))
        buf = excel.export_xlsx(current_app.config["TEMPLATE_XLSX"], rows)
        return send_file(buf, as_attachment=True,
                         download_name=f"Mintresponse_Work_Order_List_{stamp}.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument."
                                  "spreadsheetml.sheet")
    if fmt == "csv":
        return send_file(excel.export_csv(rows), as_attachment=True,
                         download_name=f"Mintresponse_Work_Orders_{stamp}.csv",
                         mimetype="text/csv")
    abort(404)


@bp.route("/template.xlsx")
def blank_template():
    return send_file(current_app.config["TEMPLATE_XLSX"], as_attachment=True,
                     download_name="Mintresponse_Work_Order_List_blank.xlsx")


@bp.route("/guide")
def guide():
    return render_template("guide.html", fields=FIELDS)


@bp.route("/api/work-orders")
def api_work_orders():
    rows = _filtered(db.all_work_orders(), request.args)
    return jsonify({"totals": calc.totals(rows), "work_orders": rows})


@bp.route("/api/calc", methods=["POST"])
def api_calc():
    """Live preview of formulas while a form is being filled in."""
    data, _errors = _form_data(request.form)
    clean = {k: (v if k not in _errors else None) for k, v in data.items()}
    out = calc.compute(clean)
    return jsonify({f.key: display(f, out.get(f.key)) for f in FIELDS if f.computed})
