from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
import json
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

app = Flask(__name__)

# ---------- Load and Save Risk Data ----------
def load_risks():
    try:
        with open("risks.json", "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return []

def save_risks(data):
    with open("risks.json", "w") as f:
        json.dump(data, f, indent=4)

# ---------- Calculate Risk Rating ----------
def calculate_rating(likelihood, impact):
    scale = {"Low": 1, "Medium": 2, "High": 3}
    rating_value = scale.get(likelihood, 0) * scale.get(impact, 0)
    if rating_value >= 6:
        return "High"
    elif rating_value >= 3:
        return "Medium"
    else:
        return "Low"

@app.route("/", methods=["GET", "POST"])
def index():
    risks = load_risks()

    if request.method == "POST":
        new_risk = {
            "Risk ID": request.form["risk_id"],
            "Risk Description": request.form["risk_description"],
            "Likelihood": request.form["likelihood"],
            "Impact": request.form["impact"],
            "Control": request.form["control"],
        }
        new_risk["Risk Rating"] = calculate_rating(new_risk["Likelihood"], new_risk["Impact"])
        risks.append(new_risk)
        save_risks(risks)
        return redirect(url_for("index"))

    return render_template("index.html", risks=risks)

# ---------- Export to Excel ----------
@app.route("/export_excel")
def export_excel():
    risks = load_risks()
    df = pd.DataFrame(risks)
    output_file = "Risk_Assessment_Report.xlsx"
    df.to_excel(output_file, index=False)
    return send_file(output_file, as_attachment=True)

# ---------- Export to PDF ----------
@app.route("/export_pdf")
def export_pdf():
    risks = load_risks()
    pdf_file = "Risk_Assessment_Report.pdf"
    doc = SimpleDocTemplate(pdf_file, pagesize=letter)
    data = [list(risks[0].keys())] + [list(r.values()) for r in risks]
    table = Table(data)
    style = TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.grey),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0,0), (-1,0), 12),
        ("BACKGROUND", (0,1), (-1,-1), colors.beige),
        ("GRID", (0,0), (-1,-1), 1, colors.black),
    ])
    table.setStyle(style)
    doc.build([table])
    return send_file(pdf_file, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)

