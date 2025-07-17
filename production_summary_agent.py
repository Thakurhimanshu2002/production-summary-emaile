import os
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime, timedelta
from email.message import EmailMessage
import smtplib

# Load environment variables
load_dotenv()

# Set yesterday's date
report_date = (datetime.today() - timedelta(days=1)).strftime('%d-%m-%Y')
yesterday = datetime.today() - timedelta(days=1)

# File path to Excel
file_path = "production.xlsx"

# Read Excel file
try:
    df = pd.read_excel(file_path)
except FileNotFoundError:
    raise Exception(f"File not found: {file_path}")

# Clean and validate columns
df.columns = df.columns.str.strip()
required_columns = ["ProdDate", "Product", "NoofBoxes", "RMCons"]
for col in required_columns:
    if col not in df.columns:
        raise Exception(f"Missing required column: {col}")

scrap_present = "scrap" in df.columns
df["ProdDate"] = pd.to_datetime(df["ProdDate"], errors='coerce')
df_day = df[df["ProdDate"].dt.date == yesterday.date()]
if df_day.empty:
    raise Exception("No production data found for yesterday.")

# Summary stats
total_punnets = int(df_day["NoofBoxes"].sum())
total_rm = float(df_day["RMCons"].sum())
total_scrap = float(df_day["scrap"].sum()) if scrap_present else 0

# Product-wise grouping
agg_dict = {"NoofBoxes": "sum", "RMCons": "sum"}
if scrap_present:
    agg_dict["scrap"] = "sum"
grouped = df_day.groupby("Product", as_index=False).agg(agg_dict)
grouped["Boxes/kg"] = grouped["NoofBoxes"] / grouped["RMCons"]
grouped = grouped.sort_values(by="NoofBoxes", ascending=False)

# Fixed-width ASCII table
col_widths = {"Product": 33, "Boxes": 10, "RM Used": 10, "Scrap": 10, "Boxes/kg": 10}
lines = []
lines.append("📊 Product-wise Performance:")
lines.append("+" + "-"*col_widths["Product"] + "+" + "-"*col_widths["Boxes"] +
             "+" + "-"*col_widths["RM Used"] + "+" + "-"*col_widths["Scrap"] +
             "+" + "-"*col_widths["Boxes/kg"] + "+")
lines.append("| {:<{}} | {:>{}} | {:>{}} | {:>{}} | {:>{}} |".format(
    "Product", col_widths["Product"],
    "Boxes", col_widths["Boxes"],
    "RM Used", col_widths["RM Used"],
    "Scrap", col_widths["Scrap"],
    "Boxes/kg", col_widths["Boxes/kg"]
))
lines.append("+" + "-"*col_widths["Product"] + "+" + "-"*col_widths["Boxes"] +
             "+" + "-"*col_widths["RM Used"] + "+" + "-"*col_widths["Scrap"] +
             "+" + "-"*col_widths["Boxes/kg"] + "+")

for _, row in grouped.iterrows():
    lines.append("| {:<{}} | {:>{},} | {:>{}.2f} | {:>{}.2f} | {:>{}.2f} |".format(
        row["Product"][:col_widths["Product"]], col_widths["Product"],
        int(row["NoofBoxes"]), col_widths["Boxes"] - 1,
        row["RMCons"], col_widths["RM Used"] - 1,
        row["scrap"] if scrap_present else 0, col_widths["Scrap"] - 1,
        row["Boxes/kg"], col_widths["Boxes/kg"] - 1
    ))

lines.append("+" + "-"*col_widths["Product"] + "+" + "-"*col_widths["Boxes"] +
             "+" + "-"*col_widths["RM Used"] + "+" + "-"*col_widths["Scrap"] +
             "+" + "-"*col_widths["Boxes/kg"] + "+")
product_summary_table = "\n".join(lines)

# Optional: HTML version
html_table_rows = ""
for _, row in grouped.iterrows():
    html_table_rows += f"""
    <tr>
        <td>{row["Product"]}</td>
        <td align='right'>{int(row["NoofBoxes"]):,}</td>
        <td align='right'>{row["RMCons"]:.2f}</td>
        <td align='right'>{row["scrap"] if scrap_present else 0:.2f}</td>
        <td align='right'>{row["Boxes/kg"]:.2f}</td>
    </tr>
    """

html_table = f"""
<h3>📊 Product-wise Performance:</h3>
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
    <thead>
        <tr>
            <th>Product</th>
            <th>Boxes</th>
            <th>RM Used</th>
            <th>Scrap</th>
            <th>Boxes/kg</th>
        </tr>
    </thead>
    <tbody>
        {html_table_rows}
    </tbody>
</table>
"""

# Efficiency message
avg_efficiency = grouped["Boxes/kg"].mean()
efficiency_msg = f"""
⚙️ Efficiency Insight:
Average material yield is {avg_efficiency:.2f} boxes per kg of raw material.
Higher Boxes/kg and lower scrap indicate good production efficiency.
"""

# Final text message
text_summary = f"""
🔷 Production Summary – {report_date}

📦 Total Output        : {total_punnets:,} punnet(s)
🧪 Raw Material Used   : {total_rm:,.2f} kg
🗑️ Scrap Generated     : {total_scrap:,.2f} kg

{product_summary_table}

{efficiency_msg.strip()}

📝 Note: Please verify all scrap entries for accuracy.
"""

# Email setup
email_from = os.getenv("EMAIL_FROM")
email_pass = os.getenv("EMAIL_PASS")
email_to = os.getenv("EMAIL_TO")

msg = EmailMessage()
msg["Subject"] = f"Production Summary - {report_date}"
msg["From"] = email_from
msg["To"] = ", ".join([email.strip() for email in email_to.split(",")])
msg.set_content(text_summary)
msg.add_alternative(f"""
<html>
<body>
    <h2>🔷 Production Summary – {report_date}</h2>
    <p><b>📦 Total Output</b>: {total_punnets:,} punnet(s)</p>
    <p><b>🧪 Raw Material Used</b>: {total_rm:,.2f} kg</p>
    <p><b>🗑️ Scrap Generated</b>: {total_scrap:,.2f} kg</p>
    {html_table}
    <p><b>⚙️ Efficiency Insight:</b><br>
    Average yield: <b>{avg_efficiency:.2f}</b> boxes/kg. Lower scrap and higher output indicate efficiency.</p>
    <p><i>📝 Note: Please verify all scrap entries for accuracy.</i></p>
</body>
</html>
""", subtype='html')

# Send email via Gmail SMTP
with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.starttls()
    server.login(email_from, email_pass)
    server.send_message(msg)

print("✅ Email sent successfully.")
