"""
EMpower Grant Dashboard — Data Processing Script
=================================================
Inputs (same folder):
  GL_Report.xlsx
  Restricted_Grant_Master_Reporting_Payments_Updated.xlsx  (new master)

Output: grants_data.json
"""

import json, os, re, subprocess, sys
from collections import defaultdict
from datetime import datetime, date

GITHUB_PUSH  = False
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
GL_FILE      = os.path.join(SCRIPT_DIR, "GL_Report.xlsx")
MASTER_FILE  = os.path.join(SCRIPT_DIR, "Restricted_Grant_Master_Reporting_Payments_Updated.xlsx")
OUTPUT_FILE  = os.path.join(SCRIPT_DIR, "grants_data.json")
GL_SHEET     = "General Ledger report (3)"

# FY periods
FY24_S=datetime(2023,7,1); FY24_E=datetime(2024,6,30)
FY25_S=datetime(2024,7,1); FY25_E=datetime(2025,6,30)
FY26_S=datetime(2025,7,1); FY26_E=datetime(2026,6,30)
FY27_S=datetime(2026,7,1); FY27_E=datetime(2027,6,30)
FY_PERIODS=[("FY24",FY24_S,FY24_E),("FY25",FY25_S,FY25_E),
            ("FY26",FY26_S,FY26_E),("FY27",FY27_S,FY27_E)]

CATEGORY_MAP={
    "7455":"Grants","7456":"Grants","7457":"Grants","7458":"Grants",
    "7459":"Grants","7460":"Grants","7461":"Grants","7462":"Grants","7463":"Grants",
    "7490":"Capacity Strengthening","7491":"Capacity Strengthening",
    "7493":"Capacity Strengthening","7495":"Capacity Strengthening",
    "7497":"Capacity Strengthening","7498":"Capacity Strengthening","7499":"Capacity Strengthening",
    "7492":"Communications",
    "7005":"Staffing","7006":"Staffing","7007":"Staffing","7008":"Staffing",
    "7105":"Staffing","7120":"Staffing","7135":"Staffing",
    "7357":"Travel","7358":"Travel","7359":"Travel","7360":"Travel",
    "7361":"Travel","7362":"Travel","7363":"Travel","7364":"Travel",
    "7365":"Travel","7366":"Travel","7367":"Travel","7368":"Travel",
    "7374":"Travel","7375":"Travel","7376":"Travel",
    "7410":"Travel","7420":"Travel",
    "7568":"Travel","7569":"Travel","7570":"Travel","7571":"Travel",
}
def get_cat(ac): return CATEGORY_MAP.get(ac,"Overhead / Other")

def parse_date(val):
    if val is None: return None
    if isinstance(val, datetime): return val
    if isinstance(val, date): return datetime(val.year,val.month,val.day)
    s=str(val).strip()
    try: return datetime.strptime(s[:10],"%Y-%m-%d")
    except: pass
    s=s.replace("-","/")
    p=s.split("/")
    if len(p)==3:
        try:
            m,d,y=int(p[0]),int(p[1]),int(p[2])
            if y<100: y+=2000
            return datetime(y,m,d)
        except: pass
    return None

def safe_float(v,default=0.0):
    if v is None: return default
    if isinstance(v,(int,float)): return float(v)
    try: return float(str(v).replace(",","").strip())
    except: return default

def fmt_date(dt): return dt.strftime("%Y-%m-%d") if dt else ""
def fy_label(p):
    for lbl,s,e in FY_PERIODS:
        if s<=p<=e: return lbl
    return "Pre-FY24" if p<FY24_S else "Future"

try: import openpyxl
except ImportError:
    subprocess.check_call([sys.executable,"-m","pip","install","openpyxl",
                           "--break-system-packages","-q"])
    import openpyxl


# ── Load new master (Restricted_Grant_Master sheet) ───────────────────────────
def load_master(path):
    print(f"Reading master: {os.path.basename(path)}")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Restricted_Grant_Master"]
    grants = {}
    for i,row in enumerate(ws.iter_rows(values_only=True)):
        if i==0: continue
        v=list(row)
        if not (v[2] and str(v[2]).strip().startswith("30-")): continue
        code=str(v[2]).strip()
        grants[code]={
            "code":code,
            "short_name":    str(v[0]).strip() if v[0] else "",
            "name":          str(v[1]).strip() if v[1] else "",
            "entity":        str(v[3]).strip() if v[3] else "",
            "start":         parse_date(v[4]),
            "end":           parse_date(v[5]),
            "budget_total":  safe_float(v[6]),
            "lead":          str(v[7]).strip() if v[7] else "",
            "rep_schedule":  str(v[8]).strip() if v[8] else "",
            "next_report_due": parse_date(v[9]),
            "payment_schedule":str(v[10]).strip() if v[10] else "",
            "next_install_due": str(v[11]).strip() if v[11] else "",
            "next_install_amt": safe_float(v[12]),
            "next_key_date":  parse_date(v[13]),
            "source_status":  str(v[14]).strip() if len(v)>14 and v[14] else "",
        }

    # Reporting_Requirements sheet
    ws2 = wb["Reporting_Requirements"]
    reporting = defaultdict(list)
    for i,row in enumerate(ws2.iter_rows(values_only=True)):
        if i==0: continue
        v=list(row)
        if not v[0]: continue
        code=str(v[0]).strip()
        reporting[code].append({
            "report_type":    str(v[2]).strip() if v[2] else "",
            "period":         str(v[3]).strip() if v[3] else "",
            "due_date":       fmt_date(parse_date(v[4])),
            "requirements":   str(v[5]).strip() if v[5] else "",
            "lead":           str(v[6]).strip() if v[6] else "",
            "status":         str(v[7]).strip() if v[7] else "",
        })

    # Payment_Installments sheet
    ws3 = wb["Payment_Installments"]
    payments = defaultdict(list)
    for i,row in enumerate(ws3.iter_rows(values_only=True)):
        if i==0: continue
        v=list(row)
        if not v[0]: continue
        code=str(v[0]).strip()
        is_next = str(v[8]).strip().lower()=="yes" if len(v)>8 and v[8] else False
        payments[code].append({
            "installment_no":  str(v[2]).strip() if v[2] else "",
            "amount":          safe_float(v[3]),
            "due_date":        str(parse_date(v[4]).strftime("%Y-%m-%d")) if parse_date(v[4]) else str(v[4]).strip() if v[4] else "",
            "condition":       str(v[5]).strip() if v[5] else "",
            "status":          str(v[6]).strip() if v[6] else "",
            "outstanding":     safe_float(v[7]),
            "is_next":         is_next,
        })

    # Needs_Review flags
    ws4 = wb["Needs_Review"]
    needs_review = {}
    for i,row in enumerate(ws4.iter_rows(values_only=True)):
        if i==0: continue
        v=list(row)
        if not v[0]: continue
        name=str(v[0]).strip()
        issue=str(v[1]).strip() if v[1] else ""
        action=str(v[2]).strip() if v[2] else ""
        # Map review items to grant codes
        for code,g in grants.items():
            if g["short_name"].lower() in name.lower() or g["name"][:10].lower() in name.lower():
                if code not in needs_review:
                    needs_review[code]=[]
                needs_review[code].append({"item":name,"issue":issue,"action":action})

    # Attach reporting/payments/flags to grants
    for code in grants:
        grants[code]["reporting_schedule"] = reporting.get(code,[])
        grants[code]["payment_installments"]= payments.get(code,[])
        grants[code]["needs_review"]        = needs_review.get(code,[])

    print(f"  Loaded {len(grants)} grants, "
          f"{sum(len(v['reporting_schedule']) for v in grants.values())} report entries, "
          f"{sum(len(v['payment_installments']) for v in grants.values())} payment installments")
    return grants


# ── Also pull budget category breakdowns from GL master (budget columns) ──────
def load_budget_cats(path):
    """Pull budget category columns from the original GL master if present."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        if "Restricted_Grant_Master" not in wb.sheetnames: return {}
        ws = wb["Restricted_Grant_Master"]
        budgets={}
        for i,row in enumerate(ws.iter_rows(values_only=True)):
            if i==0:
                hdrs=list(row)
                continue
            v=list(row)
            if not (v[2] and str(v[2]).strip().startswith("30-")): continue
            code=str(v[2]).strip()
            budgets[code]={
                "budget_grants": safe_float(v[10]),
                "budget_cs":     safe_float(v[11]),
                "budget_comms":  safe_float(v[12]),
                "budget_staff":  safe_float(v[13]),
                "budget_travel": safe_float(v[14]),
                "budget_oh":     safe_float(v[15]),
                "orig_currency": str(v[27]).strip() if len(v)>27 and v[27] else "USD",
                "agreement_status": str(v[29]).strip() if len(v)>29 and v[29] else "",
                "entity_group":  str(v[5]).strip()  if len(v)>5  and v[5]  else "",
                "purpose":       str(v[31]).strip() if len(v)>31 and v[31] else "",
                "lead_email":    str(v[24]).strip() if len(v)>24 and v[24] else "",
            }
        return budgets
    except Exception as e:
        print(f"  Note: could not load budget categories from GL master: {e}")
        return {}


# ── Load GL ───────────────────────────────────────────────────────────────────
def load_gl(path, sheet):
    print(f"Reading GL: {os.path.basename(path)}")
    wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
    if sheet not in wb.sheetnames: sheet=wb.sheetnames[0]
    ws=wb[sheet]
    rows=[]
    cur_ac=cur_an=None
    pat=re.compile(r"^(\d{4})\s*-\s*(.+?)(?:\s*\(Balance|$)")
    for i,row in enumerate(ws.iter_rows(values_only=True)):
        v=list(row)
        if i<6: continue
        if v[0] and v[1] is None and v[2] is None:
            m=pat.match(str(v[0]))
            if m: cur_ac,cur_an=m.group(1),m.group(2).strip()
            continue
        if len(v)<20: continue
        fe=str(v[10]).strip() if v[10] else ""
        if not fe.startswith("30-") or fe=="30-000": continue
        try:
            d=float(v[18]) if v[18] is not None else 0.0
            c=float(v[19]) if v[19] is not None else 0.0
        except: continue
        if d==0 and c==0: continue
        posted=parse_date(v[0])
        if posted is None: continue
        ac_p=cur_ac[0] if cur_ac else ""
        rows.append({
            "fe":fe,"ac":cur_ac,"ac_name":cur_an or "",
            "posted":posted,"debit":d,"credit":c,
            "memo":str(v[3]) if v[3] else "",
            "vendor":str(v[8]) if v[8] else "",
            "employee":str(v[6]) if v[6] else "",
            "type":"income" if ac_p=="4" else ("expense" if ac_p=="7" else "other"),
            "category":get_cat(cur_ac),
        })
    print(f"  Loaded {len(rows)} GL rows")
    return rows


# ── Build grant data ──────────────────────────────────────────────────────────
def build(grants, gl_rows, budget_cats):
    print("Building grant data...")
    result=[]
    today=datetime.today()

    for code in sorted(grants.keys()):
        g=grants[code]
        start=g["start"]
        if not start:
            print(f"  Skipping {code} — no start date"); continue

        bc=budget_cats.get(code,{})
        grant_rows=[r for r in gl_rows if r["fe"]==code]
        inc_rows  =[r for r in grant_rows if r["type"]=="income"]
        exp_rows  =[r for r in grant_rows if r["type"]=="expense" and r["posted"]>=start]

        # Income by FY
        def inc_fy(s,e): return sum(r["credit"]-r["debit"] for r in inc_rows if s<=r["posted"]<=e)
        fy_income={lbl:round(inc_fy(s,e),2) for lbl,s,e in FY_PERIODS}
        total_income=round(sum(r["credit"]-r["debit"] for r in inc_rows),2)
        income_lines=[{"date":fmt_date(r["posted"]),"memo":r["memo"][:120],
                       "vendor":r["vendor"][:80],"ac":r["ac"],
                       "amount":round(r["credit"]-r["debit"],2)}
                      for r in sorted(inc_rows,key=lambda x:x["posted"])]

        # Expenses by FY
        def exp_fy(s,e): return sum(r["debit"]-r["credit"] for r in exp_rows if s<=r["posted"]<=e)
        fy_expenses={lbl:round(exp_fy(s,e),2) for lbl,s,e in FY_PERIODS}
        total_expenses=round(sum(r["debit"]-r["credit"] for r in exp_rows),2)

        # Budget categories
        cat_budgets={
            "Grants":bc.get("budget_grants",0),
            "Capacity Strengthening":bc.get("budget_cs",0),
            "Communications":bc.get("budget_comms",0),
            "Staffing":bc.get("budget_staff",0),
            "Travel":bc.get("budget_travel",0),
            "Overhead / Other":bc.get("budget_oh",0),
        }
        cat_fy=defaultdict(lambda:defaultdict(float))
        cat_tot=defaultdict(float)
        for r in exp_rows:
            cat=r["category"]; net=r["debit"]-r["credit"]
            cat_fy[cat][fy_label(r["posted"])]+=net
            cat_tot[cat]+=net
        budget_categories=[]
        for cat,bud in cat_budgets.items():
            act=round(cat_tot.get(cat,0),2)
            fy_actuals={lbl:round(cat_fy[cat].get(lbl,0),2) for lbl,_,_ in FY_PERIODS}
            budget_categories.append({
                "category":cat,"budget":round(bud,2),
                "fy_actuals":fy_actuals,"total_actual":act,
                "variance":round(bud-act,2),
                "pct_used":round(act/bud*100,1) if bud>0 else 0,
            })

        # GL lines
        gl_lines=[{"date":fmt_date(r["posted"]),"ac":r["ac"],"ac_name":r["ac_name"][:60],
                   "category":r["category"],"memo":r["memo"][:120],
                   "vendor":r["vendor"][:80],"employee":r["employee"][:60],
                   "debit":round(r["debit"],2),"credit":round(r["credit"],2),
                   "net":round(r["debit"]-r["credit"],2),"fy":fy_label(r["posted"])}
                  for r in sorted(exp_rows,key=lambda x:x["posted"])]

        # Payment summary stats
        installs=g["payment_installments"]
        total_awarded   =round(sum(x["amount"] for x in installs),2) if installs else g["budget_total"]
        total_outstanding=round(sum(x["outstanding"] for x in installs),2) if installs else 0

        end=g["end"]
        days_remaining=(end-today).days if end else None
        grant_status="Active" if not end or today<=end else "Ended"
        budget_total=g["budget_total"]
        pct_used=round(total_expenses/budget_total*100,1) if budget_total>0 else 0

        rec={
            "code":code,"short_name":g["short_name"],"name":g["name"],
            "entity":g["entity"],
            "entity_group":bc.get("entity_group",g["entity"]),
            "start_date":fmt_date(start),"end_date":fmt_date(end),
            "days_remaining":days_remaining,"grant_status":grant_status,
            "lead":g["lead"],"lead_email":bc.get("lead_email",""),
            "orig_currency":bc.get("orig_currency","USD"),
            "budget_total":round(budget_total,2),
            "purpose":bc.get("purpose",""),
            "agreement_status":bc.get("agreement_status",""),
            "source_status":g["source_status"],
            # key dates
            "next_key_date":     fmt_date(g["next_key_date"]),
            "next_report_due":   fmt_date(g["next_report_due"]),
            "next_install_due":  g["next_install_due"],
            "next_install_amt":  round(g["next_install_amt"],2),
            # payment
            "rep_schedule_text": g["rep_schedule"],
            "payment_schedule_text": g["payment_schedule"],
            "payment_installments":  installs,
            "total_outstanding":     total_outstanding,
            # reporting
            "reporting_schedule":    g["reporting_schedule"],
            # flags
            "needs_review":          g["needs_review"],
            # income
            "fy_income":fy_income,"total_income":total_income,"income_lines":income_lines,
            # expenses
            "fy_expenses":fy_expenses,"total_expenses":total_expenses,
            "balance":round(budget_total-total_expenses,2),"pct_used":pct_used,
            # detail
            "budget_categories":budget_categories,"gl_lines":gl_lines,
        }
        result.append(rec)
        next_r=fmt_date(g["next_report_due"]) or "—"
        print(f"  {code:<10}  {g['name'][:32]:<32}  "
              f"Exp:{total_expenses:>10,.0f}  "
              f"NextReport:{next_r}  "
              f"Outstanding:{total_outstanding:>10,.0f}")
    return result


def write_json(data,path):
    out={"generated":datetime.today().strftime("%Y-%m-%d %H:%M"),
         "fy_periods":{"FY24":"01 Jul 2023 – 30 Jun 2024",
                       "FY25":"01 Jul 2024 – 30 Jun 2025",
                       "FY26":"01 Jul 2025 – 30 Jun 2026",
                       "FY27":"01 Jul 2026 – 30 Jun 2027"},
         "grants":data}
    with open(path,"w",encoding="utf-8") as f:
        json.dump(out,f,indent=2,ensure_ascii=False)
    print(f"\nOutput: {path}  ({os.path.getsize(path)/1024:.1f} KB)")


def main():
    print("="*65)
    print("EMpower Grant Dashboard — Data Processing Script")
    print("="*65+"\n")
    if not os.path.exists(GL_FILE):
        print(f"ERROR: GL not found: {GL_FILE}"); sys.exit(1)
    if not os.path.exists(MASTER_FILE):
        print(f"ERROR: Master not found: {MASTER_FILE}"); sys.exit(1)

    grants     = load_master(MASTER_FILE)
    budget_cats= load_budget_cats(GL_FILE)   # pull budget category $ from GL master
    gl_rows    = load_gl(GL_FILE, GL_SHEET)
    print()
    data = build(grants, gl_rows, budget_cats)
    print()
    write_json(data, OUTPUT_FILE)
    print("\nDone.")
    print("="*65)

if __name__=="__main__":
    main()
