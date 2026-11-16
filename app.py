"""RIMS - Railway Inspection Management System (Phase 1)
Solapur Division, Central Railway.  Streamlit + Google Sheets (no database).
"""
import hashlib, hmac, io, json, re, secrets, uuid
from datetime import date, datetime, timedelta
from urllib.parse import quote

import gspread
import pandas as pd
import streamlit as st
from docx import Document
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="RIMS", page_icon="🚆", layout="centered")

# ---------------------------------------------------------------- storage
TABS = {
    "Users": ["username", "name", "role", "salt", "pw_hash", "active"],
    "Officers": ["officer_id", "name", "designation", "department", "mobile", "email", "lang"],
    "Inspections": ["insp_id", "date", "orig_type", "orig_description", "station_section",
                    "officer_name", "status", "completed_on", "remarks", "source_file",
                    "raw_row", "created_by", "created_at"],
    "Message_Log": ["log_id", "timestamp", "insp_id", "recipient", "channel", "language",
                    "message", "sent_by"],
}


@st.cache_resource
def book():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    sh = gspread.authorize(creds).open_by_key(st.secrets["sheet_id"])
    have = {w.title for w in sh.worksheets()}
    for tab, cols in TABS.items():
        if tab not in have:
            sh.add_worksheet(tab, 2000, len(cols)).append_row(cols)
    return sh


@st.cache_data(ttl=20)
def read(tab):
    rows = book().worksheet(tab).get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(rows, columns=TABS[tab]).astype(str)


def add(tab, rows):
    book().worksheet(tab).append_rows(
        [[str(r.get(c, "")) for c in TABS[tab]] for r in rows], value_input_option="RAW")
    read.clear()


def set_cells(tab, key_col, key, **vals):
    df = read(tab)
    hit = df.index[df[key_col] == key]
    if len(hit):
        w = book().worksheet(tab)
        for col, v in vals.items():
            w.update_cell(int(hit[0]) + 2, TABS[tab].index(col) + 1, str(v))
        read.clear()


# ---------------------------------------------------------------- auth
def hp(pw, salt):
    return hashlib.pbkdf2_hmac("sha256", pw.encode(), salt.encode(), 120_000).hex()


def new_user(username, name, role, pw):
    salt = secrets.token_hex(8)
    add("Users", [dict(username=username.strip().lower(), name=name, role=role,
                       salt=salt, pw_hash=hp(pw, salt), active="yes")])


def login():
    st.title("🚆 RIMS")
    st.caption("Railway Inspection Management System · Solapur Division")
    users = read("Users")
    if users.empty:
        st.info("First run: create the Admin account.")
        with st.form("first"):
            u, n = st.text_input("Username"), st.text_input("Full name")
            p = st.text_input("Password", type="password")
            if st.form_submit_button("Create admin") and u and p:
                new_user(u, n, "Admin", p)
                st.rerun()
        return
    with st.form("login"):
        u = st.text_input("Username").strip().lower()
        p = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            row = users[(users.username == u) & (users.active == "yes")]
            if len(row) and hmac.compare_digest(hp(p, row.iloc[0].salt), row.iloc[0].pw_hash):
                st.session_state.user = row.iloc[0].to_dict()
                st.rerun()
            st.error("Invalid username or password")


# ---------------------------------------------------------------- word parsing
KEYS = {
    "date": ["date", "dt", "दिनांक"],
    "type": ["inspection", "type", "nature", "तपासणी"],
    "officer": ["officer", "inspecting", "inspector", "official", "अधिकारी"],
    "place": ["station", "section", "location", "place", "स्थानक"],
    "desc": ["description", "detail", "remark", "particular", "scope", "तपशील"],
}


def parse_docx(f):
    out = []
    for t in Document(f).tables:
        rows = [[c.text.strip() for c in r.cells] for r in t.rows]
        if len(rows) < 2:
            continue
        hdr, m = [h.lower() for h in rows[0]], {}
        for k, words in KEYS.items():
            for i, h in enumerate(hdr):
                if i not in m.values() and any(w in h for w in words):
                    m[k] = i
                    break
        for r in rows[1:]:
            if not any(r):
                continue
            g = lambda k: r[m[k]] if k in m and m[k] < len(r) else ""
            d = pd.to_datetime(g("date"), dayfirst=True, errors="coerce")
            out.append(dict(
                date=d.strftime("%Y-%m-%d") if pd.notna(d) else g("date"),
                orig_type=g("type"), orig_description=g("desc"),
                station_section=g("place"), officer_name=g("officer"),
                raw_row=json.dumps(dict(zip(rows[0], r)), ensure_ascii=False)))
    return out


# ---------------------------------------------------------------- messaging
TPL = {
    "English": ("Reminder: {type} inspection is scheduled on {date} at {place}. "
                "Inspecting officer: {officer}. Please ensure arrangements and compliance.\n"
                "- Solapur Division, Central Railway"),
}


def phone(n):
    d = re.sub(r"\D", "", n)
    return "91" + d if len(d) == 10 else d


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def overdue(df):
    d = pd.to_datetime(df["date"], errors="coerce").dt.date
    return df[(df.status != "Completed") & (d < date.today())]


# ---------------------------------------------------------------- pages
def p_dashboard():
    st.subheader("Dashboard")
    df = read("Inspections")
    if df.empty:
        st.info("No inspections yet. Upload a schedule first.")
        return
    od = overdue(df)
    c = st.columns(4)
    c[0].metric("Total", len(df))
    c[1].metric("Pending", (df.status == "Pending").sum())
    c[2].metric("Overdue", len(od))
    c[3].metric("Done", (df.status == "Completed").sum())
    st.markdown("**Officer-wise (pending)**")
    st.bar_chart(df[df.status != "Completed"].officer_name.value_counts())
    if len(od):
        st.markdown("**Overdue**")
        st.dataframe(od[["date", "orig_type", "station_section", "officer_name"]],
                     hide_index=True, use_container_width=True)


def p_upload():
    st.subheader("Upload inspection schedule (Word)")
    f = st.file_uploader("Word file (.docx)", type="docx")
    if not f:
        return
    rows = parse_docx(f)
    if not rows:
        st.error("No tables found in the document.")
        return
    st.caption("Review extracted rows. Original type/description are stored exactly as in the "
               "Word file (locked); you may correct date, place and officer.")
    df = pd.DataFrame(rows)[["date", "orig_type", "orig_description", "station_section",
                             "officer_name", "raw_row"]]
    ed = st.data_editor(df.drop(columns="raw_row"), use_container_width=True, hide_index=True,
                        disabled=["orig_type", "orig_description"], num_rows="dynamic")
    if st.button("Approve & save to register", type="primary"):
        ed = ed.reset_index(drop=True)
        recs = []
        for i, r in ed.iterrows():
            recs.append(dict(r, insp_id="I" + uuid.uuid4().hex[:8].upper(), status="Pending",
                             source_file=f.name, raw_row=df.raw_row.get(i, ""),
                             created_by=st.session_state.user["username"], created_at=now()))
        add("Inspections", recs)
        st.success(f"{len(recs)} inspections saved.")


def p_register():
    st.subheader("Inspection register")
    df = read("Inspections")
    if df.empty:
        st.info("Register is empty.")
        return
    s = st.multiselect("Status", ["Pending", "Completed", "Cancelled"], ["Pending"])
    q = st.text_input("Search (officer, station, type)")
    v = df[df.status.isin(s)]
    if q:
        v = v[v.apply(lambda r: q.lower() in " ".join(r.astype(str)).lower(), axis=1)]
    st.dataframe(v.drop(columns=["raw_row", "source_file", "created_by", "created_at"]),
                 hide_index=True, use_container_width=True)
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    st.download_button("⬇ Excel register", buf.getvalue(), "RIMS_register.xlsx")
    with st.expander("Update an inspection"):
        ids = v.insp_id.tolist()
        if ids:
            i = st.selectbox("Inspection", ids, format_func=lambda x: " | ".join(
                df[df.insp_id == x][["date", "orig_type", "station_section"]].iloc[0]))
            ns = st.selectbox("Status", ["Pending", "Completed", "Cancelled"])
            rm = st.text_area("Remarks")
            if st.button("Update"):
                set_cells("Inspections", "insp_id", i, status=ns, remarks=rm,
                          completed_on=str(date.today()) if ns == "Completed" else "")
                st.rerun()


def p_send():
    st.subheader("Send reminders")
    df, off = read("Inspections"), read("Officers")
    df = df[df.status == "Pending"]
    if df.empty:
        st.info("No pending inspections.")
        return
    days = st.slider("Due within (days)", 1, 60, 14)
    d = pd.to_datetime(df.date, errors="coerce").dt.date
    df = df[(d <= date.today() + timedelta(days=days))]
    if df.empty:
        st.info("Nothing due in this window.")
        return
    i = st.selectbox("Inspection", df.insp_id, format_func=lambda x: " | ".join(
        df[df.insp_id == x][["date", "orig_type", "station_section", "officer_name"]].iloc[0]))
    r = df[df.insp_id == i].iloc[0]
    match = off[off.name.str.lower() == r.officer_name.lower()]
    m = match.iloc[0] if len(match) else None
    lang = "English"
    msg = st.text_area("Message", TPL[lang].format(
        type=r.orig_type, date=r.date, place=r.station_section, officer=r.officer_name),
        height=150, key=f"msg{i}{lang}")
    mob = st.text_input("Mobile", m.mobile if m is not None else "")
    mail = st.text_input("Email", m.email if m is not None else "")
    c = st.columns(3)
    c[0].link_button("WhatsApp", f"https://wa.me/{phone(mob)}?text={quote(msg)}",
                     use_container_width=True)
    c[1].link_button("SMS (phone)", f"sms:{phone(mob)}?body={quote(msg)}",
                     use_container_width=True)
    c[2].link_button("Email", f"mailto:{mail}?subject={quote('Inspection reminder: ' + r.orig_type)}"
                     f"&body={quote(msg)}", use_container_width=True)
    st.caption("Google Messages Web: copy the text below (copy icon), then open it.")
    st.code(msg, language=None)
    st.link_button("Open Google Messages Web", "https://messages.google.com/web/conversations/new")
    st.divider()
    ch = st.radio("Sent via", ["WhatsApp", "SMS", "Email"], horizontal=True)
    if st.button("✔ Log as sent"):
        add("Message_Log", [dict(log_id="M" + uuid.uuid4().hex[:8].upper(), timestamp=now(),
                                 insp_id=i, recipient=mob if ch != "Email" else mail,
                                 channel=ch, language=lang, message=msg,
                                 sent_by=st.session_state.user["username"])])
        st.success("Logged.")


def p_log():
    st.subheader("Message log")
    df = read("Message_Log")
    st.dataframe(df.iloc[::-1], hide_index=True, use_container_width=True)
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    st.download_button("⬇ Excel", buf.getvalue(), "RIMS_message_log.xlsx")


def p_officers():
    st.subheader("Officer & supervisor directory")
    with st.form("off", clear_on_submit=True):
        n, des = st.text_input("Name"), st.text_input("Designation")
        dep, mob = st.text_input("Department"), st.text_input("Mobile")
        em = st.text_input("Email")
        if st.form_submit_button("Add") and n:
            add("Officers", [dict(officer_id="O" + uuid.uuid4().hex[:6].upper(), name=n,
                                  designation=des, department=dep, mobile=mob, email=em, lang="English")])
            st.rerun()
    st.dataframe(read("Officers"), hide_index=True, use_container_width=True)


def p_users():
    st.subheader("Users (Admin)")
    with st.form("usr", clear_on_submit=True):
        u, n = st.text_input("Username"), st.text_input("Full name")
        role, pw = st.selectbox("Role", ["Officer", "Admin"]), st.text_input("Password", type="password")
        if st.form_submit_button("Create user") and u and pw:
            new_user(u, n, role, pw)
            st.rerun()
    st.dataframe(read("Users")[["username", "name", "role", "active"]], hide_index=True)


# ---------------------------------------------------------------- main
if "user" not in st.session_state:
    login()
    st.stop()

user = st.session_state.user
pages = {"Dashboard": p_dashboard, "Register": p_register, "Send reminders": p_send,
         "Message log": p_log, "Officers": p_officers}
if user["role"] == "Admin":
    pages |= {"Upload schedule": p_upload, "Users": p_users}
with st.sidebar:
    st.write(f"👤 {user['name'] or user['username']} ({user['role']})")
    pg = st.radio("Menu", list(pages))
    if st.button("Logout"):
        del st.session_state.user
        st.rerun()
pages[pg]()
