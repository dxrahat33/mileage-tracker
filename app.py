import streamlit as st
import requests
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# ---------------- পেজ কনফিগারেশন ----------------
st.set_page_config(
    page_title="Scooter Mileage Tracker",
    page_icon="⚡",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ---------------- লাইট মোড ও মোবাইল UI স্টাইল ----------------
st.markdown("""
<style>
    /* মূল ব্যাকগ্রাউন্ড সাদা */
    .stApp {
        background-color: #FFFFFF;
        color: #111111;
        max-width: 480px;
        margin: 0 auto;
    }
    header, footer, #MainMenu {visibility: hidden;}

    /* উপরের অতিরিক্ত ফাঁকা জায়গা সরিয়ে মূল কনটেন্ট একেবারে ওপরে */
    .block-container {
        padding-top: 0.5rem !important;
        padding-bottom: 1rem !important;
        margin-top: 0 !important;
    }
    [data-testid="stVerticalBlock"] {
        gap: 0.35rem;
    }
    
    /* কার্ড ডিজাইন */
    .update-card {
        background: transparent;
        border: none;
        border-radius: 0;
        padding: 0;
        margin: 8px 0 18px 0;
        box-shadow: none;
    }
    .greeting-title {
        font-size: 30px;
        font-weight: 800;
        color: #111111;
        margin-bottom: 4px;
    }
    .weather-box {
        font-size: 14px;
        color: #555555;
        margin-bottom: 12px;
    }
    .stat-label {
        font-size: 13px;
        color: #555555;
        margin-bottom: 2px;
    }
    .stat-value {
        font-size: 16px;
        font-weight: 700;
        color: #111111;
        margin-bottom: 12px;
    }
    
    /* প্রাইমারি এন্ট্রি বাটন */
    div[data-testid="stButton"] > button[kind="primary"] {
        background: #111111 !important;
        color: #FFFFFF !important;
        font-weight: 800 !important;
        font-size: 17px !important;
        border: 1px solid #111111 !important;
        border-radius: 50px !important;
        padding: 14px 20px !important;
        width: 100%;
    }
    div[data-testid="stButton"] > button[kind="primary"] *,
    div[data-testid="stButton"] > button[kind="primary"] p,
    div[data-testid="stButton"] > button[kind="primary"] span {
        color: #FFFFFF !important;
    }
    
    /* সেকেন্ডারি ও ন্যাভিগেশন বাটন */
    div[data-testid="stButton"] > button[kind="secondary"] {
        background-color: #FFFFFF !important;
        color: #111111 !important;
        border: 1px solid #CCCCCC !important;
        border-radius: 14px !important;
        padding: 10px 6px !important;
        font-weight: 600 !important;
    }
    div[data-testid="stButton"] > button[kind="secondary"] *,
    div[data-testid="stButton"] > button[kind="secondary"] p,
    div[data-testid="stButton"] > button[kind="secondary"] span {
        color: #111111 !important;
    }

    /* এক্সপ্যান্ডার (রেকর্ড আইটেম) স্টাইল */
    div[data-testid="stExpander"] {
        background-color: #FFFFFF;
        border: 1px solid #D8D8D8;
        border-radius: 14px;
        margin-bottom: 10px;
    }

    /* টেক্সট, লেবেল, ইনপুট এবং ডিভাইডার */
    h1, h2, h3, h4, h5, h6, p, label, span, div {
        color: #111111;
    }
    input, textarea {
        color: #111111 !important;
        background-color: #FFFFFF !important;
    }
    hr {
        border-color: #DDDDDD;
    }
</style>
""", unsafe_allow_html=True)

# ---------------- Google Sheets কানেকশন ----------------
RECORD_HEADERS = [
    "date", "charge_time", "start_odo", "end_odo", "distance_km",
    "start_soc", "end_soc", "soc_used", "est_100_range", "km_per_pct", "completed_at"
]

def default_state():
    return {
        "user_name": "পারভেজ",
        "current_odo": 0.0,
        "current_soc": 100,
        "last_charge_time": None,
        "active_cycle": None
    }

@st.cache_resource
def get_gsheet_workbook():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    if "gcp_service_account" not in st.secrets:
        raise RuntimeError(
            "Missing [gcp_service_account] in Streamlit Secrets. "
            "Add the Google service-account fields to .streamlit/secrets.toml "
            "or Streamlit Cloud → App settings → Secrets."
        )
    if "sheet_url" not in st.secrets:
        raise RuntimeError(
            "Missing sheet_url in Streamlit Secrets. "
            "Add the Google Sheets spreadsheet URL to .streamlit/secrets.toml "
            "or Streamlit Cloud → App settings → Secrets."
        )

    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return client.open_by_url(st.secrets["sheet_url"])

def get_or_create_worksheets():
    wb = get_gsheet_workbook()
    titles = [ws.title for ws in wb.worksheets()]

    if "state" not in titles:
        ws_state = wb.add_worksheet(title="state", rows=5, cols=2)
        ws_state.update("A1", [[json.dumps(default_state(), ensure_ascii=False)]])
    else:
        ws_state = wb.worksheet("state")

    if "records" not in titles:
        ws_records = wb.add_worksheet(title="records", rows=1000, cols=15)
        ws_records.append_row(RECORD_HEADERS)
    else:
        ws_records = wb.worksheet("records")

    return ws_state, ws_records

def load_all_data():
    ws_state, ws_records = get_or_create_worksheets()
    
    raw_state = ws_state.acell("A1").value
    if raw_state:
        try:
            state = json.loads(raw_state)
        except Exception:
            state = default_state()
    else:
        state = default_state()

    rows = ws_records.get_all_records()
    state["records"] = list(reversed(rows))  # নতুন রেকর্ড সবার উপরে
    return state

def save_state_to_sheet(data):
    ws_state, _ = get_or_create_worksheets()
    state_only = {k: v for k, v in data.items() if k != "records"}
    ws_state.update("A1", [[json.dumps(state_only, ensure_ascii=False)]])

def append_record_to_sheet(record):
    _, ws_records = get_or_create_worksheets()
    row = [record[col] for col in RECORD_HEADERS]
    ws_records.append_row(row)

def reset_google_sheet():
    ws_state, ws_records = get_or_create_worksheets()
    ws_state.update("A1", [[json.dumps(default_state(), ensure_ascii=False)]])
    ws_records.clear()
    ws_records.append_row(RECORD_HEADERS)

# ---------------- লাইভ ওয়েদার ফাংশন ----------------
@st.cache_data(ttl=1800)
def get_weather(lat=23.8583, lon=90.2667):
    """সাভার, ঢাকার লাইভ ওয়েদার (Open-Meteo API - কোনো কী লাগে না)"""
    try:
        url = (
            f"https://api.open-meteo.com/v1/forecast?"
            f"latitude={lat}&longitude={lon}&current=temperature_2m,weather_code&timezone=Asia%2FDhaka"
        )
        res = requests.get(url, timeout=5).json()
        temp = round(res["current"]["temperature_2m"])
        code = res["current"]["weather_code"]
        
        weather_map = {
            0: ("☀️", "পরিষ্কার আকাশ"),
            1: ("🌤️", "প্রায় পরিষ্কার"),
            2: ("⛅", "আংশিক মেঘলা"),
            3: ("☁️", "মেঘলা আকাশ"),
            45: ("🌫️", "কুয়াশাচ্ছন্ন"),
            48: ("🌫️", "ঘন কুয়াশা"),
            51: ("🌦️", "গুঁড়ি গুঁড়ি বৃষ্টি"),
            61: ("🌧️", "হালকা বৃষ্টি"),
            63: ("🌧️", "মাঝারি বৃষ্টি"),
            65: ("⛈️", "ভারী বৃষ্টি"),
            80: ("🌦️", "বৃষ্টির সম্ভাবনা"),
            95: ("⛈️", "বজ্রসহ বৃষ্টি"),
        }
        icon, desc = weather_map.get(code, ("🌤️", "স্বাভাবিক আবহাওয়া"))
        return f"{icon} আজকের আবহাওয়া: {temp}°C, {desc}"
    except Exception:
        return "🌐 ওয়েদার আপডেট লোড করা যায়নি (অফলাইন)"

def get_greeting():
    hour = datetime.now().hour
    if 5 <= hour < 12:
        return "শুভ সকাল"
    elif 12 <= hour < 16:
        return "শুভ দুপুর"
    elif 16 <= hour < 18:
        return "শুভ বিকাল"
    elif 18 <= hour < 21:
        return "শুভ সন্ধ্যা"
    else:
        return "শুভ রাত্রি"

def time_ago_text(iso_time):
    if not iso_time:
        return "এখনো চার্জ এন্ট্রি নেই"
    try:
        dt = datetime.fromisoformat(iso_time)
        diff = datetime.now() - dt
        mins = int(diff.total_seconds() // 60)
        if mins < 1:
            return "এইমাত্র"
        if mins < 60:
            return f"{mins} মিনিট আগে"
        hours = mins // 60
        rem_mins = mins % 60
        if hours < 24:
            return f"{hours} ঘণ্টা {rem_mins} মিনিট আগে"
        days = hours // 24
        return f"{days} দিন আগে"
    except Exception:
        return "-"

def format_odo(value):
    return f"{float(value):g}"

# ---------------- সার্কুলার ব্যাটারি SVG ----------------
def battery_ring_svg(percent):
    radius = 48
    circumference = 2 * 3.1416 * radius
    offset = circumference * (1 - max(0, min(100, percent)) / 100)
    color = "#111111" if percent > 40 else ("#777777" if percent > 20 else "#444444")
    return f"""
    <div style="display:flex; justify-content:center; align-items:center; height:100%;">
        <svg width="126" height="126" viewBox="0 0 120 120">
            <circle cx="60" cy="60" r="{radius}" stroke="#E0E0E0" stroke-width="10" fill="none"/>
            <circle cx="60" cy="60" r="{radius}" stroke="{color}" stroke-width="10" fill="none"
                stroke-dasharray="{circumference}" stroke-dashoffset="{offset}"
                stroke-linecap="round" transform="rotate(-90 60 60)"/>
            <text x="60" y="67" text-anchor="middle" fill="#111111" font-size="24" font-weight="800">{percent}%</text>
        </svg>
    </div>
    """

# ---------------- স্টেট ইনিশিয়ালাইজ ----------------
if "app_data" not in st.session_state:
    try:
        st.session_state.app_data = load_all_data()
    except Exception as e:
        st.error(f"Google Sheets কানেক্ট করা যায়নি। অনুগ্রহ করে Secrets চেক করুন।\n\nError: {e}")
        st.stop()

data = st.session_state.app_data

if "page" not in st.session_state:
    st.session_state.page = "home"

# ---------------- ২-ধাপের এন্ট্রি ডায়ালগ ----------------
@st.dialog("⚡ নতুন এন্ট্রি")
def entry_dialog():
    active = data.get("active_cycle")
    
    if active:
        st.info(f"চলমান সাইকেল: চার্জ দেওয়া হয়েছিল **{active['start_soc']}%** (ODO: **{format_odo(active['start_odo'])} কি.মি.**)")
        default_tab = 1
    else:
        st.caption("প্রথমে চার্জ এন্ট্রি দিন, পরে রাইড শেষে ODO ও অবশিষ্ট চার্জ দিয়ে এন্ট্রি কমপ্লিট করুন।")
        default_tab = 0

    choice = st.radio(
        "এন্ট্রির ধরন বেছে নিন:",
        ["🔋 ১. চার্জ এন্ট্রি (শুরু)", "🛣️ ২. রাইড শেষ (ODO ও অবশিষ্ট চার্জ)"],
        index=default_tab
    )

    if choice.startswith("🔋"):
        st.markdown("##### বর্তমান চার্জ লেভেল সেট করুন")
        charge_soc = st.number_input("চার্জের পর ব্যাটারি (%)", min_value=1, max_value=100, value=100, step=1)
        start_odo = st.number_input("বর্তমান ODO (কি.মি.)", min_value=0.0, value=float(data["current_odo"]), step=0.1, format="%g")
        
        if st.button("চার্জ এন্ট্রি সেভ করুন", type="primary", use_container_width=True):
            now_iso = datetime.now().isoformat()
            data["current_soc"] = int(charge_soc)
            data["current_odo"] = round(float(start_odo), 1)
            data["last_charge_time"] = now_iso
            data["active_cycle"] = {
                "start_soc": int(charge_soc),
                "start_odo": round(float(start_odo), 1),
                "charge_time": now_iso
            }
            save_state_to_sheet(data)
            st.success("চার্জ এন্ট্রি Google Sheet-এ সেভ হয়েছে! পরের বার রাইড শেষে ODO এন্ট্রি দিন।")
            st.rerun()

    else:
        if not active:
            st.warning("আগে একটি 'চার্জ এন্ট্রি' দেওয়া নেই! অনুগ্রহ করে প্রথমে চার্জ এন্ট্রি দিন।")
            return

        st.markdown("##### রাইড শেষে বর্তমান তথ্য দিন")
        end_odo = st.number_input(
            f"বর্তমান ODO (আগের ODO: {format_odo(active['start_odo'])} কি.মি.)",
            min_value=float(active["start_odo"]),
            value=float(active["start_odo"]),
            step=0.1,
            format="%g"
        )
        end_soc = st.number_input(
            f"অবশিষ্ট ব্যাটারি % (শুরু ছিল: {active['start_soc']}%)",
            min_value=0,
            max_value=int(active["start_soc"]),
            value=max(0, int(active["start_soc"]) - 20),
            step=1
        )

        if st.button("এন্ট্রি কমপ্লিট ও রেকর্ড সেভ করুন", type="primary", use_container_width=True):
            distance = round(float(end_odo) - float(active["start_odo"]), 2)
            soc_used = int(active["start_soc"]) - int(end_soc)

            if distance <= 0:
                st.error("নতুন ODO অবশ্যই আগের ODO-এর চেয়ে বেশি হতে হবে।")
            elif soc_used <= 0:
                st.error("অবশিষ্ট ব্যাটারি অবশ্যই শুরুর চার্জের চেয়ে কম হতে হবে।")
            else:
                est_100_range = round((distance / soc_used) * 100, 1)
                km_per_pct = round(distance / soc_used, 2)
                now_iso = datetime.now().isoformat()

                record = {
                    "date": datetime.now().strftime("%d %b %Y, %I:%M %p"),
                    "charge_time": datetime.fromisoformat(active["charge_time"]).strftime("%d %b, %I:%M %p"),
                    "start_odo": active["start_odo"],
                    "end_odo": round(float(end_odo), 1),
                    "distance_km": distance,
                    "start_soc": active["start_soc"],
                    "end_soc": int(end_soc),
                    "soc_used": soc_used,
                    "est_100_range": est_100_range,
                    "km_per_pct": km_per_pct,
                    "completed_at": now_iso
                }

                append_record_to_sheet(record)
                data["records"].insert(0, record)  # নতুন রেকর্ড সবার উপরে
                data["current_odo"] = round(float(end_odo), 1)
                data["current_soc"] = int(end_soc)
                data["active_cycle"] = None  # সাইকেল কমপ্লিট
                save_state_to_sheet(data)
                st.success(f"রেকর্ড Google Sheet-এ সেভ হয়েছে! ১০০% চার্জে সম্ভাব্য রেঞ্জ: {est_100_range} কি.মি.")
                st.rerun()

# ==================== ১. হোম পেজ ====================
if st.session_state.page == "home":
    st.markdown(f'<div class="greeting-title">{get_greeting()}, {data["user_name"]}! ⚡</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="weather-box">{get_weather()}</div>', unsafe_allow_html=True)

    if data["records"]:
        avg_100_range = sum(float(r["est_100_range"]) for r in data["records"]) / len(data["records"])
        est_remaining_km = round(avg_100_range * (data["current_soc"] / 100), 1)
        range_text = f"{est_remaining_km} কি.মি."
    else:
        range_text = "ডেটা প্রয়োজন"

    last_charge_str = time_ago_text(data["last_charge_time"])
    cycle_status = "চলমান (ODO এন্ট্রি বাকি)" if data["active_cycle"] else "নতুন চার্জের জন্য প্রস্তুত"

    st.markdown('<div class="update-card">', unsafe_allow_html=True)
    st.markdown("<div style='font-size:17px; font-weight:700; margin-bottom:10px;'>আপডেটস</div>", unsafe_allow_html=True)
    
    c1, c2 = st.columns([1, 1.2])
    with c1:
        st.markdown(battery_ring_svg(data["current_soc"]), unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
            <div class="stat-label">সর্বশেষ চার্জ:</div>
            <div class="stat-value">{last_charge_str}</div>
            
            <div class="stat-label">বর্তমান ODO:</div>
            <div class="stat-value">{format_odo(data['current_odo'])} কি.মি.</div>
            
            <div class="stat-label">সম্ভাব্য রেঞ্জ ({data['current_soc']}%):</div>
            <div class="stat-value">{range_text}</div>
        """, unsafe_allow_html=True)

    st.caption(f"স্ট্যাটাস: {cycle_status}")
    st.markdown('</div>', unsafe_allow_html=True)

    btn_label = "+ রাইড শেষ এন্ট্রি দিন (ধাপ ২)" if data["active_cycle"] else "+ নতুন এন্ট্রি"
    if st.button(btn_label, type="primary", use_container_width=True):
        entry_dialog()

# ==================== ২. রেকর্ড পেজ ====================
elif st.session_state.page == "record":
    st.markdown('<div class="greeting-title" style="font-size:26px;">📋 ট্রিপ রেকর্ডস</div>', unsafe_allow_html=True)

    records = data.get("records", [])
    if not records:
        st.info("এখনো কোনো কমপ্লিট রেকর্ড নেই। হোম পেজ থেকে চার্জ এবং রাইড এন্ট্রি সম্পন্ন করুন।")
    else:
        total_km = round(sum(float(r["distance_km"]) for r in records), 1)
        avg_range = round(sum(float(r["est_100_range"]) for r in records) / len(records), 1)
        
        m1, m2 = st.columns(2)
        m1.metric("মোট চালানো হয়েছে", f"{total_km} কি.মি.")
        m2.metric("গড় রেঞ্জ (১০০%)", f"{avg_range} কি.মি.")
        st.divider()

        for rec in records:
            short_title = f"📅 {str(rec['date']).split(',')[0]}  |  🛣️ {rec['distance_km']} কি.মি.  |  🔋 -{rec['soc_used']}%  |  ⚡ {rec['est_100_range']} কি.মি."
            with st.expander(short_title, expanded=False):
                d1, d2 = st.columns(2)
                with d1:
                    st.write(f"**শুরুর ODO:** {format_odo(rec['start_odo'])} কি.মি.")
                    st.write(f"**শেষের ODO:** {format_odo(rec['end_odo'])} কি.মি.")
                    st.write(f"**মোট দূরত্ব:** {rec['distance_km']} কি.মি.")
                with d2:
                    st.write(f"**ব্যাটারি খরচ:** {rec['soc_used']}% ({rec['start_soc']}% → {rec['end_soc']}%)")
                    st.write(f"**১০০% এ সম্ভাব্য রেঞ্জ:** {rec['est_100_range']} কি.মি.")
                    st.write(f"**প্রতি ১% চার্জে:** {rec['km_per_pct']} কি.মি.")
                st.caption(f"চার্জের সময়: {rec['charge_time']}  •  শেষ হয়েছে: {rec['date']}")

# ==================== ৩. সেটিংস পেজ ====================
elif st.session_state.page == "settings":
    st.markdown('<div class="greeting-title" style="font-size:26px;">⚙️ সেটিংস</div>', unsafe_allow_html=True)

    with st.container(border=True):
        st.subheader("স্কুটার ও প্রোফাইল সেটিংস")
        new_name = st.text_input("আপনার নাম", value=data["user_name"])
        new_odo = st.number_input("বর্তমান ODO সেট করুন (কি.মি.)", min_value=0.0, value=float(data["current_odo"]), step=0.1, format="%g")
        new_soc = st.number_input("বর্তমান ব্যাটারি %", min_value=0, max_value=100, value=int(data["current_soc"]), step=1)

        if st.button("সেটিংস সেভ করুন", type="primary", use_container_width=True):
            data["user_name"] = new_name.strip() or "রাইডার"
            data["current_odo"] = round(float(new_odo), 1)
            data["current_soc"] = int(new_soc)
            save_state_to_sheet(data)
            st.success("সেটিংস Google Sheet-এ আপডেট হয়েছে!")
            st.rerun()

    st.write("")

    with st.container(border=True):
        st.subheader("⚠️ অ্যাপ রিসেট")
        st.caption("রিসেট করলে Google Sheet-এর সব ট্রিপ রেকর্ড, চলমান সাইকেল এবং ODO মুছে যাবে।")
        confirm_reset = st.checkbox("আমি নিশ্চিত যে পুরো অ্যাপ রিসেট করতে চাই")
        
        if st.button("পুরো অ্যাপ রিসেট করুন", type="secondary", use_container_width=True, disabled=not confirm_reset):
            reset_google_sheet()
            st.session_state.app_data = load_all_data()
            st.success("অ্যাপ সম্পূর্ণ রিসেট হয়েছে!")
            st.rerun()

# ==================== বটম ন্যাভিগেশন বার ====================
st.divider()
nav1, nav2, nav3 = st.columns(3)

with nav1:
    if st.button("🏠 হোম", use_container_width=True, type="primary" if st.session_state.page == "home" else "secondary"):
        st.session_state.page = "home"
        st.rerun()

with nav2:
    if st.button("📋 রেকর্ড", use_container_width=True, type="primary" if st.session_state.page == "record" else "secondary"):
        st.session_state.page = "record"
        st.rerun()

with nav3:
    if st.button("⚙️ সেটিংস", use_container_width=True, type="primary" if st.session_state.page == "settings" else "secondary"):
        st.session_state.page = "settings"
        st.rerun()