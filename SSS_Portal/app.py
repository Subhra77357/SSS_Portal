import os
import shutil
import pandas as pd
import zipfile
import tempfile
from flask import (
    Flask, render_template, request, redirect, url_for, flash,
    session, send_file, send_from_directory
)
from werkzeug.utils import secure_filename
from datetime import datetime
from io import BytesIO

# ------------------------------ CONFIG ------------------------------
app = Flask(__name__)
app.secret_key = "linuxlabs_secret_key_2025"

BASE_PATH = "/home/SUBHRAPRADEEPDAS/SSS_Portal"
UPLOAD_ROOT = os.path.join(BASE_PATH, "uploads")
TEMP_ZIP_ROOT = os.path.join(BASE_PATH, "temp_zips")
MASTER_XLSX = os.path.join(BASE_PATH, "stockist_master.xlsx")

os.makedirs(UPLOAD_ROOT, exist_ok=True)
os.makedirs(TEMP_ZIP_ROOT, exist_ok=True)

# ---------------------------- ADMIN USERS ----------------------------
ADMIN_USERS = {
    "admin@linuxlabs.com": {"password": "admin123", "division": "ALL", "role": "main"},
    "download.admin@linuxlabs.com": {"password": "download123", "division": "ALL", "role": "download"},
    "imperia.admin@linuxlabs.com": {"password": "imperia123", "division": "IMPERIA", "role": "division"},
    "infina.admin@linuxlabs.com": {"password": "infina123", "division": "INFINA", "role": "division"},
    "integra.admin@linuxlabs.com": {"password": "integra123", "division": "INTEGRA", "role": "division"},
    "dermanex.admin@linuxlabs.com": {"password": "dermanex123", "division": "DERMANEX", "role": "division"},
    "dermascience.admin@linuxlabs.com": {"password": "dermascience123", "division": "DERMASCIENCE", "role": "division"},
    "nutrimax.admin@linuxlabs.com": {"password": "nutrimax123", "division": "NUTRIMAX", "role": "division"},
    "meta.admin@linuxlabs.com": {"password": "meta123", "division": "META", "role": "division"},
}

# ----------------------------- UTILITIES -----------------------------
def ensure_master():
    cols = [
        "Division", "STATE", "RBM_HQ", "ABM_HQ", "BM_HQ",
        "Stockist_Code", "Stockist_Name", "RBM_Email", "ABM_Email", "ZBM_Email",
        "AWS_Status", "SSS_Status", "Sales_Value",
        "AWS_File", "SSS_File", "AWS_Submitted_By", "SSS_Submitted_By", "Submission_Date"
    ]
    if not os.path.exists(MASTER_XLSX):
        pd.DataFrame(columns=cols).to_excel(MASTER_XLSX, index=False)
        return
    df = pd.read_excel(MASTER_XLSX, dtype=str).fillna("")
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    save_data(df)

def load_data():
    ensure_master()
    return pd.read_excel(MASTER_XLSX, dtype=str).fillna("")

def save_data(df):
    os.makedirs(os.path.dirname(MASTER_XLSX), exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", prefix="stockist_master_")
    os.close(fd)
    try:
        df.to_excel(tmp_path, index=False, engine="openpyxl")
        shutil.move(tmp_path, MASTER_XLSX)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

def remove_file_if_exists(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass

def is_logged_in():
    return "user_email" in session or "admin" in session

def current_submitter_email():
    return session.get("user_email") or session.get("admin_email") or ""

def ensure_upload_folder(division, state, kind):
    folder = os.path.join(UPLOAD_ROOT, division, state, kind)
    os.makedirs(folder, exist_ok=True)
    return folder
# --------------------------- FILE SAVE HELPER ---------------------------
def save_file(file, division, state, kind, stockist_name, stockist_code):
    """Save uploaded file inside uploads/<division>/<state>/<kind>/ with clean filename."""
    ext = os.path.splitext(file.filename)[1]
    clean_name = secure_filename(f"{stockist_name}_{stockist_code}_{kind}{ext}")
    folder = ensure_upload_folder(division, state, kind)
    path = os.path.join(folder, clean_name)
    file.save(path)
    return clean_name


# ------------------------------ ROUTES ------------------------------
@app.route("/")
def home():
    if "user_email" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form["email"].strip().lower()
        df = load_data()
        if (
            email in df["RBM_Email"].str.lower().values
            or email in df["ABM_Email"].str.lower().values
            or email in df["ZBM_Email"].str.lower().values
        ):
            session.clear()
            session["user_email"] = email
            flash("Login successful!", "success")
            return redirect(url_for("dashboard"))
        flash("Unauthorized email!", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    flash("Logged out successfully.", "info")
    return redirect(url_for("login"))

@app.route("/dashboard")
def dashboard():
    if "user_email" not in session:
        return redirect(url_for("login"))
    email = session["user_email"]
    df = load_data()
    df = df[
        (df["RBM_Email"].str.lower() == email)
        | (df["ABM_Email"].str.lower() == email)
        | (df["ZBM_Email"].str.lower() == email)
    ]
    rows = df.to_dict(orient="records")
    summary = {
        "total": len(rows),
        "aws_submitted": sum(r["AWS_Status"] == "Submitted" for r in rows),
        "aws_pending": sum(r["AWS_Status"] != "Submitted" for r in rows),
        "sss_submitted": sum(r["SSS_Status"] == "Submitted" for r in rows),
        "sss_pending": sum(r["SSS_Status"] != "Submitted" for r in rows),
    }
    return render_template("dashboard.html", rows=rows, summary=summary, email=email)

@app.route("/admin", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        email = request.form["email"].strip().lower()
        password = request.form["password"].strip()
        user = ADMIN_USERS.get(email)
        if user and user["password"] == password:
            session.clear()
            session["admin"] = True
            session["admin_email"] = email
            session["admin_division"] = user["division"]
            session["admin_role"] = user["role"]
            flash("Admin login successful!", "success")
            if user["role"] == "download":
                return redirect(url_for("admin_downloads_page"))
            return redirect(url_for("admin_dashboard"))
        flash("Invalid credentials!", "danger")
    return render_template("admin_login.html")

@app.route("/admin_logout")
def admin_logout():
    session.clear()
    flash("Admin logged out.", "info")
    return redirect(url_for("admin_login"))

@app.route("/upload_aws", methods=["POST"])
def upload_aws():
    submitter = current_submitter_email()
    if not submitter:
        flash("Please login first.", "danger")
        return redirect(url_for("login"))
    stockist_code = request.form.get("stockist_code")
    sales_value = request.form.get("sales_value", "").strip()
    df = load_data()
    if stockist_code not in df["Stockist_Code"].values:
        flash("Stockist not found!", "warning")
        return redirect(request.referrer or url_for("dashboard"))
    row = df[df["Stockist_Code"] == stockist_code].iloc[0]
    file = request.files.get("aws_files")
    if not file or file.filename == "":
        flash("No file selected!", "warning")
        return redirect(request.referrer or url_for("dashboard"))
    filename = save_file(file, row["Division"], row["STATE"], "AWS", row["Stockist_Name"], stockist_code)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_sales = sales_value or row.get("Sales_Value", "")
    df.loc[df["Stockist_Code"] == stockist_code, ["Sales_Value", "AWS_Status", "AWS_File", "AWS_Submitted_By", "Submission_Date"]] = [
        new_sales, "Submitted", filename, submitter, now
    ]
    save_data(df)
    flash(f"AWS uploaded for {row['Stockist_Name']}", "success")
    if "admin" in session:
        return redirect(url_for("admin_dashboard"))
    return redirect(url_for("dashboard"))

@app.route("/upload_sss", methods=["POST"])
def upload_sss():
    submitter = current_submitter_email()
    if not submitter:
        flash("Please login first.", "danger")
        return redirect(url_for("login"))
    stockist_code = request.form.get("stockist_code")
    sales_value = request.form.get("sales_value", "").strip()
    df = load_data()
    if stockist_code not in df["Stockist_Code"].values:
        flash("Stockist not found!", "warning")
        return redirect(request.referrer or url_for("dashboard"))
    row = df[df["Stockist_Code"] == stockist_code].iloc[0]
    file = request.files.get("sss_files")
    if not file or file.filename == "":
        flash("No file selected!", "warning")
        return redirect(request.referrer or url_for("dashboard"))
    filename = save_file(file, row["Division"], row["STATE"], "SSS", row["Stockist_Name"], stockist_code)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_sales = sales_value or row.get("Sales_Value", "")
    df.loc[df["Stockist_Code"] == stockist_code, ["Sales_Value", "SSS_Status", "SSS_File", "SSS_Submitted_By", "Submission_Date"]] = [
        new_sales, "Submitted", filename, submitter, now
    ]
    save_data(df)
    flash(f"SSS uploaded for {row['Stockist_Name']}", "success")
    if "admin" in session:
        return redirect(url_for("admin_dashboard"))
    return redirect(url_for("dashboard"))

@app.route("/admin_update_sales", methods=["POST"])
def admin_update_sales():
    if "admin" not in session:
        flash("Please log in as Admin.", "danger")
        return redirect(url_for("admin_login"))
    stockist_code = request.form.get("stockist_code")
    sales_value = request.form.get("sales_value", "").strip()
    if not stockist_code:
        flash("Invalid stockist code.", "warning")
        return redirect(url_for("admin_dashboard"))
    df = load_data()
    if stockist_code not in df["Stockist_Code"].values:
        flash("Stockist not found.", "warning")
        return redirect(url_for("admin_dashboard"))
    if sales_value == "":
        df.loc[df["Stockist_Code"] == stockist_code, ["Sales_Value"]] = [""]
        message = "Sales Value deleted successfully."
    else:
        df.loc[df["Stockist_Code"] == stockist_code, ["Sales_Value", "Submission_Date"]] = [
            sales_value, datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
        message = "Sales Value updated successfully!"
    save_data(df)
    flash(message, "success")
    return redirect(url_for("admin_dashboard"))

@app.route("/admin_delete/<stockist_code>/<kind>")
def admin_delete(stockist_code, kind):
    if "admin" not in session:
        return redirect(url_for("admin_login"))
    df = load_data()
    if stockist_code not in df["Stockist_Code"].values:
        flash("Stockist not found!", "warning")
        return redirect(url_for("admin_dashboard"))
    row = df[df["Stockist_Code"] == stockist_code].iloc[0]
    division, state = row["Division"], row["STATE"]
    if kind == "AWS" and row.get("AWS_File", ""):
        path = os.path.join(UPLOAD_ROOT, division, state, "AWS", row["AWS_File"])
        if os.path.exists(path):
            os.remove(path)
        df.loc[df["Stockist_Code"] == stockist_code, ["AWS_Status", "AWS_File"]] = ["", ""]
    elif kind == "SSS" and row.get("SSS_File", ""):
        path = os.path.join(UPLOAD_ROOT, division, state, "SSS", row["SSS_File"])
        if os.path.exists(path):
            os.remove(path)
        df.loc[df["Stockist_Code"] == stockist_code, ["SSS_Status", "SSS_File"]] = ["", ""]
    elif kind == "Sales":
        df.loc[df["Stockist_Code"] == stockist_code, ["Sales_Value"]] = [""]
    df.loc[df["Stockist_Code"] == stockist_code, "Submission_Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_data(df)
    flash(f"{kind} deleted and Excel updated!", "info")
    return redirect(url_for("admin_dashboard"))

@app.route("/admin_dashboard")
def admin_dashboard():
    if "admin" not in session:
        return redirect(url_for("admin_login"))
    role = session.get("admin_role")
    division = session.get("admin_division")
    if role == "download":
        flash("You only have access to the download page.", "warning")
        return redirect(url_for("admin_downloads_page"))
    df = load_data()
    if role == "division":
        df = df[df["Division"].str.upper() == division.upper()]
    rows = df.to_dict(orient="records")
    summary = {
        "total": len(df),
        "aws_submitted": int((df["AWS_Status"] == "Submitted").sum()),
        "aws_pending": int((df["AWS_Status"] != "Submitted").sum()),
        "sss_submitted": int((df["SSS_Status"] == "Submitted").sum()),
        "sss_pending": int((df["SSS_Status"] != "Submitted").sum()),
    }
    divisions = [{"name": d} for d in sorted(load_data()["Division"].unique())]
    return render_template("admin_dashboard.html", rows=rows, summary=summary, divisions=divisions, role=role, admin_division=division)

@app.route("/admin_download_stockist_master")
@app.route("/admin_download_stockist_master/<division>")
def admin_download_stockist_master(division=None):
    if "admin" not in session:
        return redirect(url_for("admin_login"))
    df = load_data()
    role = session.get("admin_role")
    session_div = session.get("admin_division")
    if role == "division":
        df = df[df["Division"].str.upper() == session_div.upper()]
        filename = f"{session_div}_Stockist_Master.xlsx"
    elif role in ["main", "download"] and division:
        df = df[df["Division"].str.upper() == division.upper()]
        filename = f"{division}_Stockist_Master.xlsx"
    else:
        filename = "All_Divisions_Stockist_Master.xlsx"
    out = BytesIO()
    df.to_excel(out, index=False, engine="openpyxl")
    out.seek(0)
    return send_file(out, download_name=filename, as_attachment=True)

@app.route("/admin_download_all")
def admin_download_all():
    if "admin" not in session:
        return redirect(url_for("admin_login"))
    if session.get("admin_role") not in ["main", "download"]:
        flash("Access denied.", "danger")
        return redirect(url_for("admin_downloads_page"))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"All_Divisions_Uploads_{timestamp}.zip"
    zip_path = os.path.join(TEMP_ZIP_ROOT, zip_name)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(UPLOAD_ROOT):
            for f in files:
                file_path = os.path.join(root, f)
                rel_path = os.path.relpath(file_path, UPLOAD_ROOT)
                zf.write(file_path, rel_path)
    try:
        return_data = send_file(zip_path, as_attachment=True, download_name=zip_name)
    finally:
        remove_file_if_exists(zip_path)
    return return_data

@app.route("/admin_download_division_all_states/<division>")
def admin_download_division_all_states(division):
    if "admin" not in session:
        return redirect(url_for("admin_login"))
    role = session.get("admin_role")
    session_div = session.get("admin_division")
    division = secure_filename(division)
    if role == "division" and session_div.upper() != division.upper():
        flash("Access denied for this division.", "danger")
        return redirect(url_for("admin_downloads_page"))
    div_path = os.path.join(UPLOAD_ROOT, division)
    if not os.path.isdir(div_path):
        flash("Division folder not found.", "warning")
        return redirect(url_for("admin_downloads_page"))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"{division}_All_States_{timestamp}.zip"
    zip_path = os.path.join(TEMP_ZIP_ROOT, zip_name)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(div_path):
            for f in files:
                rel = os.path.relpath(os.path.join(root, f), UPLOAD_ROOT)
                zf.write(os.path.join(root, f), rel)
    try:
        return_data = send_file(zip_path, as_attachment=True, download_name=zip_name)
    finally:
        remove_file_if_exists(zip_path)
    return return_data

@app.route("/admin_downloads_state/<division>/<state>/<kind>")
def admin_downloads_state(division, state, kind):
    if "admin" not in session:
        return redirect(url_for("admin_login"))
    role = session.get("admin_role")
    session_div = session.get("admin_division")
    division, state, kind = map(secure_filename, [division, state, kind])
    if role == "division" and session_div.upper() != division.upper():
        flash("Access denied.", "danger")
        return redirect(url_for("admin_downloads_page"))
    folder = os.path.join(UPLOAD_ROOT, division, state, kind)
    if not os.path.isdir(folder):
        flash("Folder not found.", "warning")
        return redirect(url_for("admin_downloads_page"))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zip_name = f"{division}_{state}_{kind}_{timestamp}.zip"
    zip_path = os.path.join(TEMP_ZIP_ROOT, zip_name)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder):
            for f in files:
                rel = os.path.relpath(os.path.join(root, f), UPLOAD_ROOT)
                zf.write(os.path.join(root, f), rel)
    try:
        return_data = send_file(zip_path, as_attachment=True, download_name=zip_name)
    finally:
        remove_file_if_exists(zip_path)
    return return_data

@app.route("/admin_downloads_page")
def admin_downloads_page():
    if "admin" not in session:
        return redirect(url_for("admin_login"))
    role = session.get("admin_role")
    division = session.get("admin_division")
    divisions_info = []
    def get_last_updated(folder_path):
        if not os.path.isdir(folder_path):
            return ""
        latest = max(
            (os.path.getmtime(os.path.join(folder_path, f))
             for f in os.listdir(folder_path)
             if os.path.isfile(os.path.join(folder_path, f))),
            default=None
        )
        return datetime.fromtimestamp(latest).strftime("%Y-%m-%d %I:%M %p") if latest else ""
    if role in ["main", "download"]:
        for d in sorted(os.listdir(UPLOAD_ROOT)):
            div_path = os.path.join(UPLOAD_ROOT, d)
            if os.path.isdir(div_path):
                states = []
                for s in sorted(os.listdir(div_path)):
                    state_path = os.path.join(div_path, s)
                    if os.path.isdir(state_path):
                        states.append({"name": s, "last_updated": get_last_updated(state_path)})
                divisions_info.append({"name": d, "states": states})
    else:
        div_path = os.path.join(UPLOAD_ROOT, division)
        if os.path.isdir(div_path):
            states = []
            for s in sorted(os.listdir(div_path)):
                state_path = os.path.join(div_path, s)
                if os.path.isdir(state_path):
                    states.append({"name": s, "last_updated": get_last_updated(state_path)})
            divisions_info.append({"name": division, "states": states})
    return render_template("admin_downloads_page.html", divisions=divisions_info, role=role, admin_division=division)

@app.route("/serve_upload/<division>/<state>/<kind>/<filename>")
def serve_upload(division, state, kind, filename):
    if not is_logged_in():
        return redirect(url_for("login"))
    folder = os.path.join(UPLOAD_ROOT, secure_filename(division), secure_filename(state), secure_filename(kind))
    return send_from_directory(folder, secure_filename(filename), as_attachment=False)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
