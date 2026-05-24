"""
Ind AS 116 – Lease Liability Working
Streamlit App

Run:   streamlit run IndAS116_Streamlit.py
Needs: pip install streamlit openpyxl python-dateutil pandas
"""

import math, io
from datetime import date, datetime
from dateutil.relativedelta import relativedelta

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─── page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Ind AS 116 – Lease Liability",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background:#F7F7F5; }
.block-container { padding:1.5rem 2.2rem 3rem; max-width:1440px; }

/* title bar */
.ttl { background:#1A1A1A; border-radius:10px; padding:14px 22px; margin-bottom:1.2rem; }
.ttl h1 { color:#fff; font-size:20px; font-weight:700; margin:0; }
.ttl p  { color:#9CA3AF; font-size:12px; margin:2px 0 0; }

/* card */
.card { background:#fff; border:1px solid #E5E7EB; border-radius:10px;
        padding:1.1rem 1.4rem; margin-bottom:1rem; }
.clabel { font-size:10.5px; font-weight:700; color:#6B7280;
          text-transform:uppercase; letter-spacing:.07em; margin-bottom:.7rem; }

/* metric cards */
.mcard { background:#fff; border:1px solid #E5E7EB; border-radius:8px;
         padding:12px 16px; }
.mcard .ml { font-size:11px; color:#6B7280; margin-bottom:3px; }
.mcard .mv { font-size:19px; font-weight:700; color:#111827; }
.mcard .mv.g { color:#1D9E75; }

/* info callout */
.callout { background:#EFF6FF; border-left:3px solid #3B82F6;
           border-radius:0 6px 6px 0; padding:9px 13px;
           font-size:12px; color:#1E40AF; margin:6px 0 10px; }

/* override streamlit widgets */
label { font-size:12px !important; font-weight:600 !important; color:#374151 !important; }

/* calculate button */
div[data-testid="stButton"] > button {
    background:#1A1A1A !important; color:#fff !important;
    border:none !important; border-radius:7px !important;
    font-weight:700 !important; font-size:13px !important;
    padding:9px 0 !important; width:100%;
}
div[data-testid="stButton"] > button:hover { background:#374151 !important; }

/* download button */
div[data-testid="stDownloadButton"] > button {
    background:#1D9E75 !important; color:#fff !important;
    border:none !important; border-radius:7px !important;
    font-weight:700 !important; font-size:13px !important;
    padding:9px 0 !important; width:100%;
}
div[data-testid="stDownloadButton"] > button:hover { background:#178a64 !important; }

/* tabs */
.stTabs [data-baseweb="tab-list"] {
    gap:2px; background:#F3F4F6; border-radius:8px 8px 0 0; padding:4px 4px 0;
}
.stTabs [data-baseweb="tab"] {
    font-size:13px; font-weight:500; border-radius:6px 6px 0 0;
    padding:8px 20px; color:#6B7280;
}
.stTabs [aria-selected="true"] {
    background:#fff !important; color:#1A1A1A !important;
    border-bottom:2px solid #1A1A1A !important;
}

div[data-testid="stDataFrame"] { border-radius:8px; overflow:hidden; }
hr { border:none; border-top:1px solid #E5E7EB; margin:1rem 0; }
</style>
""", unsafe_allow_html=True)


# ─── helpers ──────────────────────────────────────────────────────────────────
def add_months(dt: date, n: int) -> date:
    return dt + relativedelta(months=n)

def parse_date(s: str) -> date:
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y", "%d %b %Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Cannot parse '{s}'. Use DD-MM-YYYY.")

def inr(n: float) -> str:
    return f"₹ {n:,.2f}"


# ─── calculation engine ────────────────────────────────────────────────────────
def compute_schedule(from_date, to_date, base_payment, annual_rate,
                     months_override, esc_rate, esc_every_n):
    monthly_rate = annual_rate / 100 / 12

    if months_override and months_override > 0:
        num_months = int(months_override)
    else:
        rd = relativedelta(to_date, from_date)
        num_months = rd.years * 12 + rd.months
        if num_months <= 0:
            raise ValueError("Lease end date must be after start date.")

    n = max(1, int(esc_every_n))

    payments = []
    for i in range(1, num_months + 1):
        periods = (i - 1) // n
        esc     = math.pow(1 + esc_rate / 100, periods)
        payments.append(round(base_payment * esc, 2))

    pv       = sum(payments[i] / math.pow(1 + monthly_rate, i + 1) for i in range(num_months))
    rou_depr = pv / num_months

    rows, open_liab = [], pv
    for i in range(num_months):
        mp       = payments[i]
        df       = 1 / math.pow(1 + monthly_rate, i + 1)
        interest = open_liab * monthly_rate
        cl       = open_liab + interest - mp
        cl       = 0.0 if abs(cl) < 0.005 else cl
        rou      = pv - rou_depr * (i + 1)
        rou      = 0.0 if abs(rou) < 0.005 else rou
        rows.append({
            "sno": i + 1,
            "date": add_months(from_date, i + 1),
            "mp": mp, "df": df,
            "pv": round(mp * df, 2),
            "open": round(open_liab, 2),
            "int":  round(interest, 2),
            "close": round(cl, 2),
            "depr":  round(rou_depr, 2),
            "rou":   round(rou, 2),
        })
        open_liab = cl

    return rows, pv, monthly_rate, num_months, rou_depr


# ─── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(schedule, info) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = info["name"][:31]
    ws.sheet_view.showGridLines = False

    def W(row, col, val, bold=False, align="left", fmt=None,
          bg=None, fg="1A1A1A", size=10, wrap=False):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = Font(name="Calibri", bold=bold, size=size, color=fg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
        if bg:  c.fill          = PatternFill("solid", fgColor=bg)
        if fmt: c.number_format = fmt
        return c

    def MW(row, c1, c2, val, **kw):
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
        W(row, c1, val, **kw)

    # title
    ws.row_dimensions[1].height = 30
    MW(1,1,11,"IND AS 116 – LEASE LIABILITY WORKING",
       bold=True,align="center",size=14,bg="1A1A1A",fg="FFFFFF")
    ws.row_dimensions[2].height = 22
    MW(2,1,11, info["name"],
       bold=True,align="center",size=11,bg="1D9E75",fg="FFFFFF")

    # info block
    meta = [
        ("Lease Term (Years)",                  f"{info['num_months']/12:.2f}"),
        ("Beginning From",                      info["from_date"].strftime("%d-%b-%Y")),
        ("Ending To",                           info["to_date"].strftime("%d-%b-%Y")),
        ("No. of Months",                       info["num_months"]),
        ("Annual Discount Rate",                info["annual_rate"] / 100),
        ("Monthly Discount Rate",               info["monthly_rate"]),
        ("Base Monthly Payment (Rs.)",          info["base_payment"]),
        ("Escalation Rate (%)",                 info["esc_rate"]),
        ("Escalation Frequency (every N months)", info["esc_every_n"]),
        ("PV of Lease Payments – Initial Lease Liability", info["pv"]),
    ]
    IR = 4
    for idx, (lbl, val) in enumerate(meta):
        r = IR + idx
        ws.row_dimensions[r].height = 15
        W(r, 1, lbl, bold=True, fg="374151", size=9)
        is_pct = isinstance(val, float) and idx in (4, 5)
        last   = idx == len(meta) - 1
        cell   = W(r, 2, val,
                   fmt="0.00%" if is_pct else "#,##0.00" if isinstance(val, float) else None,
                   fg="111827", size=9, bold=last)
        if last:
            cell.font = Font(name="Calibri", bold=True, size=10, color="1D9E75")

    # schedule header
    HDR = IR + len(meta) + 2
    ws.row_dimensions[HDR].height = 34
    COLS = [
        ("S.No",                    4),
        ("Month",                  13),
        ("Lease\nPayment\n(Rs.)",  16),
        ("Discount\nFactor",       14),
        ("Present\nValue\n(Rs.)",  16),
        ("Opening\nLiability\n(Rs.)", 19),
        ("Interest\nExpense\n(Rs.)",  16),
        ("Lease\nPayment\n(Rs.)",  16),
        ("Closing\nLiability\n(Rs.)", 19),
        ("Depreciation\n(Rs.)",    15),
        ("ROU Asset\n(Rs.)",       15),
    ]
    for ci, (hdr, wid) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = wid
        c = W(HDR, ci, hdr, bold=True, align="center",
               bg="1A1A1A", fg="FFFFFF", size=9, wrap=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # data rows
    NF = "#,##0.00"
    for ri, row in enumerate(schedule):
        r  = HDR + 1 + ri
        bg = "F3F4F6" if ri % 2 else None
        ws.row_dimensions[r].height = 14

        def wr(col, val, fmt=None, align="right", _r=r, _bg=bg):
            return W(_r, col, val, fmt=fmt, align=align, bg=_bg, size=9)

        wr(1,  row["sno"],   align="center")
        wr(2,  row["date"],  fmt="DD-MMM-YYYY", align="center")
        wr(3,  row["mp"],    fmt=NF)
        wr(4,  row["df"],    fmt="0.00000000")
        wr(5,  row["pv"],    fmt=NF)
        wr(6,  row["open"],  fmt=NF)
        wr(7,  row["int"],   fmt=NF)
        wr(8,  row["mp"],    fmt=NF)
        wr(9,  row["close"], fmt=NF)
        wr(10, row["depr"],  fmt=NF)
        wr(11, row["rou"],   fmt=NF)

    # year-1 journal summary
    SUM = HDR + len(schedule) + 3
    ws.row_dimensions[SUM].height = 18
    MW(SUM,1,11,"YEAR 1 – JOURNAL ENTRY SUMMARY",
       bold=True,align="center",bg="374151",fg="FFFFFF",size=10)

    yr1     = schedule[:min(12, len(schedule))]
    yr1_d   = sum(r["depr"] for r in yr1)
    yr1_i   = sum(r["int"]  for r in yr1)
    yr1_p   = sum(r["mp"]   for r in yr1)

    jrows = [
        ("Depreciation A/c",                "Dr", yr1_d, ""),
        ("   To ROU Asset A/c",             "Cr", "",    yr1_d),
        ("", "", "", ""),
        ("Interest on Lease Liability A/c", "Dr", yr1_i, ""),
        ("   To Lease Liability A/c",       "Cr", "",    yr1_i),
        ("", "", "", ""),
        ("Lease Liability A/c",             "Dr", yr1_p, ""),
        ("   To Bank / Cash A/c",           "Cr", "",    yr1_p),
    ]
    for ji, (part, dc, dr, cr) in enumerate(jrows):
        r = SUM + 2 + ji
        ws.row_dimensions[r].height = 14
        W(r, 1, part, size=9, fg="111827")
        W(r, 2, dc,   size=9, fg="6B7280", bold=True, align="center")
        W(r, 3, dr if dr != "" else None, fmt=NF, align="right", size=9)
        W(r, 4, cr if cr != "" else None, fmt=NF, align="right", size=9)

    ws.freeze_panes = f"A{HDR + 1}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─── UI ───────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="ttl">
  <h1>📋 Ind AS 116 – Lease Liability Working</h1>
  <p>Amortisation schedule &nbsp;·&nbsp; ROU asset &nbsp;·&nbsp;
     Journal entries &nbsp;·&nbsp; Excel export</p>
</div>
""", unsafe_allow_html=True)

# ── input form ────────────────────────────────────────────────────────────────
with st.container():
    st.markdown('<div class="clabel">Lease Details</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        lease_name = st.text_input("Lease Name",
            placeholder="e.g. Gachibowli Office – 3rd Floor")
    with c2:
        lease_from = st.text_input("Lease Start Date", placeholder="DD-MM-YYYY")
    with c3:
        lease_to   = st.text_input("Lease End Date",   placeholder="DD-MM-YYYY")

    c4, c5, c6 = st.columns(3)
    with c4:
        monthly_payment = st.number_input("Monthly Lease Payment (₹)",
            min_value=0.0, value=0.0, step=1000.0, format="%.2f")
    with c5:
        annual_rate = st.number_input("Annual Discount Rate (%)",
            min_value=0.01, max_value=50.0, value=11.0, step=0.01, format="%.2f")
    with c6:
        months_override = st.number_input(
            "No. of Months (0 = auto from dates)",
            min_value=0, value=0, step=1)

    st.markdown("---")
    st.markdown('<div class="clabel">Escalation Settings</div>', unsafe_allow_html=True)

    ce1, ce2 = st.columns(2)
    with ce1:
        esc_rate = st.number_input(
            "Escalation Rate (%)",
            min_value=0.0, max_value=100.0, value=0.0, step=0.01, format="%.2f",
            help="% increase applied to the lease payment at each escalation step.")
    with ce2:
        esc_every_n = st.number_input(
            "Escalation Every N Months",
            min_value=1, max_value=360, value=12, step=1,
            help=(
                "Payment increases after every N months.\n"
                "12 = annual  |  6 = semi-annual  |  3 = quarterly  |  1 = monthly"
            ),
        )

    if esc_rate > 0:
        example_after = round(monthly_payment * (1 + esc_rate / 100), 2) if monthly_payment else "—"
        st.markdown(
            f'<div class="callout">💡 Lease payment increases by <b>{esc_rate:.2f}%</b> '
            f'every <b>{esc_every_n} month(s)</b>. '
            f'{"Base ₹" + f"{monthly_payment:,.0f}" + " → after " + str(esc_every_n) + " months: ₹" + f"{example_after:,.2f}" if monthly_payment else ""}'
            f'</div>',
            unsafe_allow_html=True)

    st.markdown("")
    btn_col, _ = st.columns([1, 4])
    with btn_col:
        do_calc = st.button("⚡  Calculate Schedule", use_container_width=True)

# ── run calculation ────────────────────────────────────────────────────────────
if do_calc:
    errs = []
    if not lease_name.strip():
        errs.append("Lease name is required.")
    if monthly_payment <= 0:
        errs.append("Monthly payment must be > 0.")
    if not lease_from.strip() and months_override == 0:
        errs.append("Enter lease start date.")

    from_date = to_date = None
    if lease_from.strip():
        try:    from_date = parse_date(lease_from)
        except ValueError as e: errs.append(str(e))
    else:
        from_date = date.today()

    if lease_to.strip():
        try:    to_date = parse_date(lease_to)
        except ValueError as e: errs.append(str(e))

    if months_override == 0 and to_date is None:
        errs.append("Enter lease end date or set number of months.")

    for e in errs:
        st.error(e)
    if errs:
        st.stop()

    eff_to = to_date if to_date else add_months(from_date, int(months_override))

    try:
        sched, pv, mr, nm, rou_d = compute_schedule(
            from_date, eff_to, monthly_payment,
            annual_rate, months_override, esc_rate, esc_every_n)
    except Exception as e:
        st.error(f"Calculation error: {e}")
        st.stop()

    st.session_state.update({
        "sched": sched, "pv": pv, "mr": mr, "nm": nm, "rou_d": rou_d,
        "info": {
            "name": lease_name.strip() or "Lease",
            "from_date": from_date, "to_date": eff_to,
            "num_months": nm, "annual_rate": annual_rate,
            "monthly_rate": mr, "pv": pv,
            "base_payment": monthly_payment,
            "esc_rate": esc_rate, "esc_every_n": esc_every_n,
            "rou_depr": rou_d,
        }
    })

# ── results ────────────────────────────────────────────────────────────────────
if "sched" in st.session_state:
    sched = st.session_state["sched"]
    pv    = st.session_state["pv"]
    mr    = st.session_state["mr"]
    nm    = st.session_state["nm"]
    rou_d = st.session_state["rou_d"]
    info  = st.session_state["info"]

    st.markdown("---")

    # metric strip
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Lease Term",            f"{nm} months")
    m2.metric("Annual Discount Rate",  f"{info['annual_rate']:.2f}%")
    m3.metric("Monthly Discount Rate", f"{mr*100:.4f}%")
    m4.metric("Initial Lease Liability", inr(pv))
    m5.metric("Monthly ROU Depreciation", inr(rou_d))

    st.markdown("")

    # tabs
    tab1, tab2, tab3 = st.tabs([
        "  📊 Amortisation Schedule  ",
        "  📖 Journal Entries (Yr 1)  ",
        "  📋 Payment Step Summary  ",
    ])

    # ── tab 1: schedule ───────────────────────────────────────────────────────
    with tab1:
        rows_display = []
        for r in sched:
            rows_display.append({
                "S.No":                  r["sno"],
                "Month":                 r["date"].strftime("%d-%b-%Y"),
                "Lease Payment (₹)":     f"{r['mp']:,.2f}",
                "Discount Factor":       f"{r['df']:.8f}",
                "Present Value (₹)":     f"{r['pv']:,.2f}",
                "Opening Liability (₹)": f"{r['open']:,.2f}",
                "Interest Expense (₹)":  f"{r['int']:,.2f}",
                "Lease Payment (₹) ":    f"{r['mp']:,.2f}",
                "Closing Liability (₹)": f"{r['close']:,.2f}",
                "Depreciation (₹)":      f"{r['depr']:,.2f}",
                "ROU Asset (₹)":         f"{r['rou']:,.2f}",
            })

        st.dataframe(pd.DataFrame(rows_display),
                     use_container_width=True, hide_index=True, height=430)

        total_int  = sum(r["int"] for r in sched)
        total_pay  = sum(r["mp"]  for r in sched)
        total_depr = sum(r["depr"] for r in sched)
        st.caption(
            f"Rows: {nm}  ·  Total lease payments: {inr(total_pay)}  ·  "
            f"Total interest expense: {inr(total_int)}  ·  "
            f"Total depreciation: {inr(total_depr)}")

    # ── tab 2: journal entries ────────────────────────────────────────────────
    with tab2:
        yr1    = sched[:min(12, nm)]
        yr1_d  = sum(r["depr"] for r in yr1)
        yr1_i  = sum(r["int"]  for r in yr1)
        yr1_p  = sum(r["mp"]   for r in yr1)

        st.markdown("**Initial Recognition (Day 1)**")
        st.dataframe(pd.DataFrame([
            {"Particulars": "ROU Asset A/c",
             "Dr (₹)": f"{pv:,.2f}", "Cr (₹)": "—"},
            {"Particulars": "   To Lease Liability A/c",
             "Dr (₹)": "—", "Cr (₹)": f"{pv:,.2f}"},
        ]), hide_index=True, use_container_width=True, height=100)

        st.markdown("**Monthly entries – Year 1**")
        jrows = []
        for r in yr1:
            d = r["date"].strftime("%d-%b-%Y")
            jrows += [
                {"Date": d,  "Particulars": "Interest on Lease Liability A/c  Dr",
                 "Dr (₹)": f"{r['int']:,.2f}", "Cr (₹)": "—", "Narration": "Interest accrual"},
                {"Date": "",  "Particulars": "   To Lease Liability A/c",
                 "Dr (₹)": "—", "Cr (₹)": f"{r['int']:,.2f}", "Narration": ""},
                {"Date": d,  "Particulars": "Lease Liability A/c  Dr",
                 "Dr (₹)": f"{r['mp']:,.2f}", "Cr (₹)": "—", "Narration": "Lease payment"},
                {"Date": "",  "Particulars": "   To Bank / Cash A/c",
                 "Dr (₹)": "—", "Cr (₹)": f"{r['mp']:,.2f}", "Narration": ""},
                {"Date": d,  "Particulars": "Depreciation A/c  Dr",
                 "Dr (₹)": f"{r['depr']:,.2f}", "Cr (₹)": "—", "Narration": "ROU depreciation"},
                {"Date": "",  "Particulars": "   To ROU Asset A/c",
                 "Dr (₹)": "—", "Cr (₹)": f"{r['depr']:,.2f}", "Narration": ""},
                {"Date": "·", "Particulars": "·", "Dr (₹)": "", "Cr (₹)": "", "Narration": ""},
            ]
        st.dataframe(pd.DataFrame(jrows),
                     hide_index=True, use_container_width=True, height=430)

        st.markdown("**Year 1 Summary**")
        last_yr1 = sched[min(11, nm - 1)]
        st.dataframe(pd.DataFrame([
            {"Particulars": "Total Depreciation (Yr 1)",    "Amount (₹)": f"{yr1_d:,.2f}"},
            {"Particulars": "Total Interest Expense (Yr 1)","Amount (₹)": f"{yr1_i:,.2f}"},
            {"Particulars": "Total Lease Payments (Yr 1)",  "Amount (₹)": f"{yr1_p:,.2f}"},
            {"Particulars": "Closing Lease Liability",      "Amount (₹)": f"{last_yr1['close']:,.2f}"},
            {"Particulars": "Closing ROU Asset",            "Amount (₹)": f"{last_yr1['rou']:,.2f}"},
        ]), hide_index=True, use_container_width=True, height=215)

    # ── tab 3: payment steps ──────────────────────────────────────────────────
    with tab3:
        st.markdown("**Payment steps across the lease term**")
        steps, cur_pay, s_month = [], None, 1
        for i, r in enumerate(sched):
            if r["mp"] != cur_pay:
                if cur_pay is not None:
                    steps.append({"From Month": s_month, "To Month": i,
                                  "No. of Months": i - s_month + 1,
                                  "Monthly Payment (₹)": f"{cur_pay:,.2f}"})
                cur_pay, s_month = r["mp"], i + 1
        steps.append({"From Month": s_month, "To Month": nm,
                      "No. of Months": nm - s_month + 1,
                      "Monthly Payment (₹)": f"{cur_pay:,.2f}"})
        st.dataframe(pd.DataFrame(steps),
                     hide_index=True, use_container_width=True)
        st.caption(
            f"Payment escalates by {info['esc_rate']:.2f}% every "
            f"{info['esc_every_n']} month(s). "
            f"Total payment steps: {len(steps)}.")

    # ── download ──────────────────────────────────────────────────────────────
    st.markdown("---")
    dcol, icol = st.columns([1, 3])
    with dcol:
        xl_bytes = build_excel(sched, info)
        fname    = f"IndAS116_{info['name'].replace(' ','_')}.xlsx"
        st.download_button(
            label="⬇️  Download Excel",
            data=xl_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with icol:
        st.caption(
            f"Excel includes: title block · lease info · "
            f"full {nm}-month amortisation schedule · "
            f"Year 1 journal entry summary · "
            f"freeze panes · number formatting")
