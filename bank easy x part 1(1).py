"""
Part 1: Bank Statement -> Tally Import XML
Client uploads PDF / Excel / CSV bank statement -> app parses it -> shows preview
-> downloads Tally XML (vouchers posted against the bank ledger).

Run:
    pip install -r requirements.txt
    python app.py
Open http://127.0.0.1:5000
"""
import io
import re
from datetime import datetime
from xml.sax.saxutils import escape

import pandas as pd
import pdfplumber
from flask import Flask, request, render_template_string, Response

app = Flask(__name__)

# ----------------------------------------------------------------------------
# 1. PARSING
# ----------------------------------------------------------------------------
DATE_KEYS = ["date", "txn date", "transaction date", "value date", "tran date"]
NARR_KEYS = ["narration", "description", "particulars", "remarks", "details", "transaction remarks"]
DEBIT_KEYS = ["debit", "withdrawal", "withdrawals", "dr", "withdrawal amt", "debit amount"]
CREDIT_KEYS = ["credit", "deposit", "deposits", "cr", "deposit amt", "credit amount"]
AMOUNT_KEYS = ["amount"]
TYPE_KEYS = ["dr/cr", "type", "cr/dr"]


def to_number(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    s = re.sub(r"[^\d.\-]", "", str(v).replace(",", ""))
    try:
        return abs(float(s)) if s not in ("", "-", ".") else 0.0
    except ValueError:
        return 0.0


def to_date(v):
    if isinstance(v, (datetime, pd.Timestamp)):
        return pd.Timestamp(v).to_pydatetime()
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%d %b %Y",
                "%d-%b-%Y", "%d-%b-%y", "%Y-%m-%d", "%d.%m.%Y", "%d %b %y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def find_col(cols, keys):
    for c in cols:
        if str(c).strip().lower() in keys:
            return c
    for c in cols:  # partial match
        if any(k in str(c).strip().lower() for k in keys):
            return c
    return None


def normalise(df):
    """Find the header row, map columns, return list of transaction dicts."""
    df = df.dropna(how="all").reset_index(drop=True)
    # header row = first row that contains a date-like and a narration-like word
    for i in range(min(len(df), 40)):
        row = [str(x).strip().lower() for x in df.iloc[i].tolist()]
        if any(r in DATE_KEYS or "date" in r for r in row) and \
           any(any(k in r for k in NARR_KEYS) for r in row):
            df.columns = [str(x).strip() for x in df.iloc[i].tolist()]
            df = df.iloc[i + 1:].reset_index(drop=True)
            break

    cols = list(df.columns)
    c_date = find_col(cols, DATE_KEYS)
    c_narr = find_col(cols, NARR_KEYS)
    c_dr = find_col(cols, DEBIT_KEYS)
    c_cr = find_col(cols, CREDIT_KEYS)
    c_amt = find_col(cols, AMOUNT_KEYS)
    c_type = find_col(cols, TYPE_KEYS)
    if not c_date or not c_narr:
        raise ValueError("Could not find Date / Narration columns in the statement.")

    txns = []
    for _, r in df.iterrows():
        d = to_date(r[c_date])
        if not d:
            continue
        debit = credit = 0.0
        if c_dr and c_cr:
            debit, credit = to_number(r[c_dr]), to_number(r[c_cr])
        elif c_amt:
            amt = to_number(r[c_amt])
            t = str(r[c_type]).strip().lower() if c_type else ""
            if t.startswith("cr"):
                credit = amt
            else:
                debit = amt
        if debit == 0 and credit == 0:
            continue
        txns.append({
            "date": d,
            "narration": re.sub(r"\s+", " ", str(r[c_narr])).strip(),
            "debit": debit,    # money OUT of bank  -> Payment
            "credit": credit,  # money INTO bank    -> Receipt
        })
    return txns


def parse_excel_csv(file):
    name = file.filename.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file, header=None, dtype=str)
    else:
        df = pd.read_excel(file, header=None, dtype=str)
    return normalise(df)


def parse_pdf(file):
    frames = []
    with pdfplumber.open(io.BytesIO(file.read())) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if table and len(table) > 1:
                    frames.append(pd.DataFrame(table))
    if not frames:
        raise ValueError("No tables found in PDF (scanned PDFs need OCR - Part 2).")
    return normalise(pd.concat(frames, ignore_index=True))


# ----------------------------------------------------------------------------
# 2. TALLY XML
# ----------------------------------------------------------------------------
def build_tally_xml(txns, company, bank_ledger, suspense_ledger):
    out = []
    out.append(f"""<ENVELOPE>
 <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
 <BODY><IMPORTDATA>
  <REQUESTDESC>
   <REPORTNAME>Vouchers</REPORTNAME>
   <STATICVARIABLES><SVCURRENTCOMPANY>{escape(company)}</SVCURRENTCOMPANY></STATICVARIABLES>
  </REQUESTDESC>
  <REQUESTDATA>""")
    for t in txns:
        is_receipt = t["credit"] > 0
        amt = t["credit"] if is_receipt else t["debit"]
        vtype = "Receipt" if is_receipt else "Payment"
        date = t["date"].strftime("%Y%m%d")
        narr = escape(t["narration"])
        # Payment : Dr Party / Cr Bank      Receipt : Dr Bank / Cr Party
        bank_pos = "Yes" if is_receipt else "No"
        party_pos = "No" if is_receipt else "Yes"
        bank_amt = -amt if is_receipt else amt
        party_amt = amt if is_receipt else -amt
        out.append(f"""   <TALLYMESSAGE xmlns:UDF="TallyUDF">
    <VOUCHER VCHTYPE="{vtype}" ACTION="Create" OBJVIEW="Accounting Voucher View">
     <DATE>{date}</DATE>
     <VOUCHERTYPENAME>{vtype}</VOUCHERTYPENAME>
     <NARRATION>{narr}</NARRATION>
     <PARTYLEDGERNAME>{escape(suspense_ledger)}</PARTYLEDGERNAME>
     <ALLLEDGERENTRIES.LIST>
      <LEDGERNAME>{escape(suspense_ledger)}</LEDGERNAME>
      <ISDEEMEDPOSITIVE>{party_pos}</ISDEEMEDPOSITIVE>
      <AMOUNT>{party_amt:.2f}</AMOUNT>
     </ALLLEDGERENTRIES.LIST>
     <ALLLEDGERENTRIES.LIST>
      <LEDGERNAME>{escape(bank_ledger)}</LEDGERNAME>
      <ISDEEMEDPOSITIVE>{bank_pos}</ISDEEMEDPOSITIVE>
      <AMOUNT>{bank_amt:.2f}</AMOUNT>
     </ALLLEDGERENTRIES.LIST>
    </VOUCHER>
   </TALLYMESSAGE>""")
    out.append("  </REQUESTDATA>\n </IMPORTDATA></BODY>\n</ENVELOPE>")
    return "\n".join(out)


# ----------------------------------------------------------------------------
# 3. WEB UI
# ----------------------------------------------------------------------------
PAGE = """
<!doctype html><title>Bank Statement to Tally</title>
<style>
 body{font-family:system-ui;max-width:900px;margin:30px auto;padding:0 16px}
 input,button{padding:8px;margin:4px 0}
 label{display:block;margin-top:10px;font-weight:600}
 table{border-collapse:collapse;width:100%;margin-top:16px}
 td,th{border:1px solid #ccc;padding:6px;font-size:14px}
 .err{color:#b00020}
</style>
<h2>Bank Statement &rarr; Tally</h2>
<form method="post" enctype="multipart/form-data">
 <label>Statement file (PDF / XLSX / XLS / CSV)</label>
 <input type="file" name="file" required>
 <label>Tally company name (exact)</label>
 <input name="company" value="{{company}}" required>
 <label>Bank ledger name in Tally</label>
 <input name="bank" value="{{bank}}" required>
 <label>Ledger for unmatched entries</label>
 <input name="suspense" value="{{suspense}}" required>
 <br><button>Upload &amp; Preview</button>
</form>
{% if error %}<p class="err">{{error}}</p>{% endif %}
{% if txns %}
 <h3>{{txns|length}} transactions found</h3>
 <form method="post" action="/download" enctype="multipart/form-data">
  <input type="hidden" name="company" value="{{company}}">
  <input type="hidden" name="bank" value="{{bank}}">
  <input type="hidden" name="suspense" value="{{suspense}}">
  <input type="hidden" name="payload" value="{{payload}}">
  <button>Download Tally XML</button>
 </form>
 <table><tr><th>Date</th><th>Narration</th><th>Debit (Out)</th><th>Credit (In)</th></tr>
 {% for t in txns %}<tr><td>{{t.date.strftime('%d-%m-%Y')}}</td><td>{{t.narration}}</td>
 <td>{{'%.2f'|format(t.debit) if t.debit else ''}}</td>
 <td>{{'%.2f'|format(t.credit) if t.credit else ''}}</td></tr>{% endfor %}
 </table>
{% endif %}
"""

import json


def pack(txns):
    return json.dumps([{**t, "date": t["date"].strftime("%Y-%m-%d")} for t in txns])


def unpack(s):
    return [{**t, "date": datetime.strptime(t["date"], "%Y-%m-%d")} for t in json.loads(s)]


@app.route("/", methods=["GET", "POST"])
def index():
    ctx = dict(company="My Company", bank="HDFC Bank", suspense="Suspense A/c",
               txns=None, payload="", error=None)
    if request.method == "POST":
        ctx.update(company=request.form["company"], bank=request.form["bank"],
                   suspense=request.form["suspense"])
        f = request.files["file"]
        try:
            if f.filename.lower().endswith(".pdf"):
                txns = parse_pdf(f)
            else:
                txns = parse_excel_csv(f)
            if not txns:
                raise ValueError("No transactions found.")
            ctx["txns"] = txns
            ctx["payload"] = pack(txns).replace('"', "&quot;")
        except Exception as e:
            ctx["error"] = f"Error: {e}"
    return render_template_string(PAGE, **ctx)


@app.route("/download", methods=["POST"])
def download():
    txns = unpack(request.form["payload"])
    xml = build_tally_xml(txns, request.form["company"], request.form["bank"],
                          request.form["suspense"])
    return Response(xml, mimetype="application/xml",
                    headers={"Content-Disposition": "attachment; filename=tally_import.xml"})


if __name__ == "__main__":
    app.run(debug=True)
