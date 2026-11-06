import streamlit as st
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
import math
import re

TAIPEI = ZoneInfo("Asia/Taipei")

# ==========================================
# 0. 網頁基本設定 & 快取功能 (🌟 解決 409 衝突的核心)
# ==========================================
APP_NAME = "中創園區 餐廳管理平台"
st.set_page_config(page_title=APP_NAME, page_icon="🍱", layout="wide")

APP_CSS = """<style>
.block-container{padding-top:3.6rem;max-width:1400px}   /* 留空間給上方工具列，日期標籤才不會被切到 */
.hero{display:flex;gap:14px;align-items:center;padding:6px 0 2px}
.hero-icon{font-size:40px;width:64px;height:64px;border-radius:16px;background:linear-gradient(135deg,#ffedd5,#fed7aa);
  display:flex;align-items:center;justify-content:center}
.hero-title{font-size:30px;font-weight:800;letter-spacing:.5px;color:#1f2937}
.hero-sub{font-size:14px;color:#6b7280;margin-top:2px}
.hero-sub a{color:#c2410c;text-decoration:none;font-weight:600}
.scards{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:14px 0 6px}
.scard{border-radius:14px;padding:12px 16px;border:1px solid #e5e7eb;background:#fff;box-shadow:0 1px 2px rgba(0,0,0,.04);
  border-left:6px solid #9ca3af}
.scard.ok{border-left-color:#16a34a;background:#f0fdf4}
.scard.todo{border-left-color:#f59e0b;background:#fffbeb}
.scard.info{border-left-color:#ea580c;background:#fff7ed}
.s-time{font-size:12px;color:#6b7280;font-family:ui-monospace,monospace}
.s-title{font-size:14px;font-weight:700;color:#374151;margin-top:2px}
.s-value{font-size:24px;font-weight:800;color:#111827;margin:2px 0}
.s-value a{color:#c2410c;text-decoration:none;font-size:18px}
.s-note{font-size:12px;color:#6b7280}
.stTabs [data-baseweb="tab-list"]{gap:6px;border-bottom:2px solid #f3f4f6}
.stTabs [data-baseweb="tab"]{font-size:16px;font-weight:600;padding:10px 18px;border-radius:10px 10px 0 0}
.stTabs [aria-selected="true"]{background:#fff7ed}
h3{color:#1f2937}
@media (max-width:900px){.scards{grid-template-columns:repeat(2,minmax(0,1fr))}.hero-title{font-size:22px}}
</style>"""

# 🌟 新增快取機制：把下載下來的資料暫存 600 秒 (10 分鐘)
# 這樣就算您在網頁上狂點數字，也不會重複去向 Google 要資料而被阻擋！
@st.cache_data(ttl=600)
def load_csv_data(url):
    return pd.read_csv(url)

# ==========================================
# 🌟 Google Sheets「CTIC_Access_Log」連線（讀卡機統計、回寫現金）
# 需在 Streamlit Cloud 的 Settings → Secrets 貼上 [gcp_service_account]（google_key.json 內容）
# 沒設定時不影響原本功能，只是改回手動輸入
# ==========================================
ACCESS_LOG_SHEET_ID = "1VI2Iw9EACSn-v7PjTB4wCG0cLeySg97tQ9GESWdilik"
COUNTS_SHEET = "每日刷卡統計"   # 筆電下午一條龍寫入
TREND_SHEET = "趨勢紀錄"
ORDER_SHEET = "團膳訂單"        # 每天一列：早上送出的份數＋加量原因、主菜、換菜、下午結算
ORDER_HEADERS = ["日期", "戰情室第一版", "實際正常數量", "換菜折合", "方便素", "用餐人數", "總訂餐份數", "寫入時間", "加購主菜",
                 "臨時加量", "加量原因", "主菜", "主菜類型", "換菜明細",
                 "供餐調整", "總供餐數", "實際加購主菜", "現場付現", "便當盒", "環電", "結算寫入時間"]
AM_COLS = ("A", "N")            # 早上按鈕寫 A~N
PM_COLS = ("O", "U")            # 下午按鈕寫 O~U
MAIN_TYPES = ["整塊（雞腿、排骨等，一人一塊）", "散肉（肉片、肉絲，容易多夾）"]
# 趨勢紀錄的發佈 CSV（不需要金鑰），用來算戰情室的 08:30 第一版
TREND_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSv8X_NJDPjSVJTEkFjdBTSJKjDnZ18eyf6yBdTBq2eVT8TB6_wB-UDUUVUoF5eUJrLSnDsqpu7OhbS/pub?gid=1825976142&single=true&output=csv"
DASHBOARD_URL = "https://tsaitzuchien-star.github.io/CTIC_Dining_Prediction/"
FALLBACK_BASE = {0: 132, 1: 128, 2: 119, 3: 126, 4: 108}   # 趨勢紀錄資料不足時的星期基準量
BASE_LOOKBACK = 8

CARD_SHEET = "新增卡片申請"     # App 送出，筆電 process_card_requests.py 處理
CARD_HEADERS = ["申請時間", "員工編號", "姓名", "公司", "卡號", "狀態", "處理時間", "備註"]

@st.cache_data(ttl=60)
def load_card_requests():
    try:
        return open_access_log().worksheet(CARD_SHEET).get_all_records()
    except Exception as e:
        if type(e).__name__ == "WorksheetNotFound":
            return []
        return f"{type(e).__name__}: {e}"

FORGOT_SHEET = "忘記帶卡補登"   # App 送出，筆電 process_card_requests.py 寫進 Vendor_Access
FORGOT_HEADERS = ["申請時間", "用餐日期", "員工編號", "姓名", "公司", "狀態", "處理時間", "備註"]

RENAME_SHEET = "姓名更正"       # 與補登同欄位：姓名欄放「正確姓名」，備註放「原姓名：xxx」
CANCEL_SHEET = "刷卡取消"      # 刷卡後沒吃到：筆電把那筆從 Vendor_Access 移到 Vendor_Access_Cancelled

def forgot_ws(sheet=FORGOT_SHEET):
    book = open_access_log()
    try:
        return book.worksheet(sheet)
    except Exception:
        ws = book.add_worksheet(title=sheet, rows=1000, cols=len(FORGOT_HEADERS))
        ws.update(values=[FORGOT_HEADERS], range_name="A1")
        return ws

@st.cache_data(ttl=60)
def load_forgot_requests(sheet=FORGOT_SHEET):
    try:
        return open_access_log().worksheet(sheet).get_all_records()
    except Exception as e:
        return [] if type(e).__name__ == "WorksheetNotFound" else f"{type(e).__name__}: {e}"

@st.cache_data(ttl=600)
def load_vendor_people():
    """員工與常客主檔裡的廠商同仁（員工編號 CTIC- 開頭），補登用"""
    try:
        rows = open_access_log().worksheet("員工與常客主檔").get_all_records()
    except Exception as e:
        return f"{type(e).__name__}: {e}"
    norm = lambda c: "環電" if "環電" in c else (c or "未分類")   # 主檔裡「環電」「環電股份有限公司」併成一家
    return sorted(({"員工編號": str(r["員工編號"]), "姓名": str(r["姓名"]).strip(), "公司": norm(str(r["公司"]).strip()),
                    "卡號": str(r.get("卡號", "")).strip()}
                   for r in rows if str(r.get("員工編號", "")).startswith("CTIC")), key=lambda p: p["姓名"])

FIXED_PLATE = ["高健維", "林盛郎"]   # 現場餐盤每天固定用餐（請假時在畫面上拿掉）
VEG_LOG_SHEET = "素食訂購紀錄"   # 每天每人一列，統計素食訂購（含臨時加訂）
VEG_LOG_HEADERS = ["日期", "姓名", "飯量", "名單內"]

@st.cache_data(ttl=120)
def load_veg_log(date_str):
    """當天記錄的方便素名單（重開網頁時帶回用）"""
    try:
        return [r for r in open_access_log().worksheet(VEG_LOG_SHEET).get_all_records() if str(r.get("日期")) == date_str]
    except Exception:
        return []

def split_names(s):
    import re
    return [n.strip() for n in re.split(r"[、,，\s]+", s or "") if n.strip()]

def write_veg_log(date_str, people, listed):
    """把當天的方便素名單寫進「素食訂購紀錄」，同一天再按一次會整天重寫"""
    book = open_access_log()
    try:
        ws = book.worksheet(VEG_LOG_SHEET)
    except Exception:
        ws = book.add_worksheet(title=VEG_LOG_SHEET, rows=2000, cols=len(VEG_LOG_HEADERS))
        ws.update(values=[VEG_LOG_HEADERS], range_name="A1")
    dates = ws.col_values(1)
    for r in sorted((i + 1 for i, d in enumerate(dates) if d == date_str), reverse=True):
        ws.delete_rows(r)
    if people:
        ws.append_rows([[date_str, n, rice, "是" if n in listed else "否（臨時）"] for n, rice in people],
                       value_input_option="RAW")
    temp = sum(n not in listed for n, _ in people)
    return f"已記錄素食 {len(people)} 位" + (f"（臨時加訂 {temp} 位）" if temp else "")

def has_gs_secret():
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False

def open_access_log():
    import gspread
    gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
    return gc.open_by_key(ACCESS_LOG_SHEET_ID)

@st.cache_data(ttl=300)
def load_daily_counts(date_str):
    """回傳當天卡機統計 dict；查不到回傳 None，錯誤回傳字串"""
    try:
        rows = open_access_log().worksheet(COUNTS_SHEET).get_all_records()
    except Exception as e:
        return f"{type(e).__name__}: {e}"
    for r in rows:
        if str(r.get("日期", "")).strip() == date_str:
            return r
    return None

def first_version_base(day):
    """戰情室 08:30 第一版：近 8 次同星期「實際用餐＋現金」中位數（與戰情室 index.html 相同）"""
    if day.weekday() > 4:
        return None, "假日"
    try:
        df = load_csv_data(TREND_CSV_URL)
        d = pd.to_datetime(df["日期"].astype(str).str.replace("/", "-"), errors="coerce")
        act = pd.to_numeric(df["實際用餐"], errors="coerce")
        cash = pd.to_numeric(df.get("現金付款", 0), errors="coerce").fillna(0)
        t = pd.DataFrame({"d": d, "v": act + cash, "a": act}).dropna()
        t = t[(t["d"].dt.date < day) & (t["d"].dt.weekday == day.weekday()) & (t["a"] > 0)]
        vals = t.sort_values("d")["v"].tail(BASE_LOOKBACK)
        if len(vals) >= 4:
            return int(math.floor(vals.median() + 0.5)), f"近 {len(vals)} 次同星期實際用餐（含現金）中位數"
    except Exception as e:
        return FALLBACK_BASE[day.weekday()], f"固定星期基準量（趨勢紀錄讀取失敗：{e}）"
    return FALLBACK_BASE[day.weekday()], "固定星期基準量（趨勢資料不足）"

def order_ws():
    """「團膳訂單」分頁；舊表只有 A~I 欄，第一次用時自動補上新欄位標題"""
    ws = open_access_log().worksheet(ORDER_SHEET)
    if ws.col_count < len(ORDER_HEADERS):
        ws.add_cols(len(ORDER_HEADERS) - ws.col_count)
    if ws.row_values(1) != ORDER_HEADERS:
        ws.update(values=[ORDER_HEADERS], range_name="A1")
    return ws

def write_order_part(date_str, values, cols, what):
    """把一段欄位寫進當天那列（沒有就新增），同一天再按一次就覆蓋"""
    ws = order_ws()
    dates = ws.col_values(1)
    if date_str in dates:
        r = dates.index(date_str) + 1
    else:
        r = len(dates) + 1
        ws.update(values=[[date_str]], range_name=f"A{r}", value_input_option="USER_ENTERED")
    ws.update(values=[values], range_name=f"{cols[0]}{r}:{cols[1]}{r}", value_input_option="USER_ENTERED")
    return f"已記錄 {date_str} 的{what}"

@st.cache_data(ttl=120)
def load_today_status(date_str):
    """今日進度：早上送出記錄（團膳訂單）與現金寫回（趨勢紀錄 E 欄）；讀不到回傳空 dict"""
    status = {}
    try:
        book = open_access_log()
        for r in book.worksheet(ORDER_SHEET).get_all_records():
            if str(r.get("日期", "")).strip() == date_str:
                status["order"] = r
        for r in book.worksheet(TREND_SHEET).get_all_values()[1:]:
            if r and r[0].strip().replace("/", "-") == date_str and len(r) > 4 and str(r[4]).strip():
                status["cash"] = r[4]
    except Exception as e:
        status["error"] = f"{type(e).__name__}: {e}"
    return status

def write_cash_to_trend(date_str, cash):
    """把現金份數寫進「趨勢紀錄」的 現金付款 欄（E 欄）；沒有這天就新增一列"""
    ws = open_access_log().worksheet(TREND_SHEET)
    dates = ws.col_values(1)
    if date_str in dates:
        row = dates.index(date_str) + 1
        ws.update(values=[[cash]], range_name=f"E{row}", value_input_option="USER_ENTERED")
        return f"已更新 {date_str} 的現金付款為 {cash} 份"
    ws.append_row([date_str, "", "", "", cash], value_input_option="USER_ENTERED")
    return f"已新增 {date_str}，現金付款 {cash} 份"

# ==========================================
# 1. 日期與雲端資料網址
# ==========================================
current_time = datetime.now(TAIPEI)  # Streamlit Cloud 主機是 UTC，要換成台灣時間

MENU_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT9QdhFOdM2cp7FI1qu4VNRvwOF6mHDJZ7OP0iYTu2shMiF5PrZI3lUzyP436KyBV3uv49akqBytF47/pub?output=csv"
VEG_CSV_URL = "https://docs.google.com/spreadsheets/d/1dGsbEe6aCJo0gexj5Xo2gdTmQ_oA6E4VNIdmEHZDGZM/export?format=csv&gid=1496853361"
RISK_CSV_URL = "https://docs.google.com/spreadsheets/d/1dGsbEe6aCJo0gexj5Xo2gdTmQ_oA6E4VNIdmEHZDGZM/export?format=csv&gid=2090477701"
NAMES_CSV_URL = "https://docs.google.com/spreadsheets/d/1dGsbEe6aCJo0gexj5Xo2gdTmQ_oA6E4VNIdmEHZDGZM/export?format=csv&gid=777001"  # 外帶餐盤名單

EDIT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1dGsbEe6aCJo0gexj5Xo2gdTmQ_oA6E4VNIdmEHZDGZM/edit"
ITRI_CANTEEN_URL = "https://restorg.itri.org.tw/index.aspx"
CTIC_LIFE_URL = "https://sites.google.com/d/1QxzEC936SJmN9fFzheP6cbl43eveS2FY/p/1VcccyLG4yToXJgO56nnF93Q-xuLg233w/edit?pli=1"

weekdays_ch = ["一", "二", "三", "四", "五", "六", "日"]

def bullet_list(names):
    return "\n    - " + "\n    - ".join(names) if names else "無"

# ==========================================
# 2. 標題、日期、今日進度
# ==========================================
st.markdown(APP_CSS, unsafe_allow_html=True)
col_title, col_date = st.columns([4, 1], vertical_alignment="bottom")
with col_title:
    st.markdown(f"""<div class="hero"><div class="hero-icon">🍱</div><div>
<div class="hero-title">{APP_NAME}</div>
<div class="hero-sub">訂餐 · 備餐 · 結算 · 卡片管理｜今日份數由 <a href="{DASHBOARD_URL}" target="_blank">📊 用餐預測戰情室</a> 自動帶入</div>
</div></div>""", unsafe_allow_html=True)
with col_date:
    selected_date = st.date_input("📅 日期", value=current_time.date())
today_str = selected_date.strftime("%Y-%m-%d")
backup_week = weekdays_ch[selected_date.weekday()]
display_date = f"{today_str} ({backup_week})"
is_friday = selected_date.weekday() == 4

status = load_today_status(today_str) if has_gs_secret() else {}
order_row = status.get("order")
if has_gs_secret():
    counts_top = load_daily_counts(today_str)
    def card(time_, title, value, note, ok):
        cls = "ok" if ok is True else ("todo" if ok is False else "info")
        return (f'<div class="scard {cls}"><div class="s-time">{time_}</div><div class="s-title">{title}</div>'
                f'<div class="s-value">{value}</div><div class="s-note">{note}</div></div>')
    cards = [
        card("08:30", "早上訂餐",
             f"{order_row.get('總訂餐份數', '')} 份" if order_row else "還沒記錄",
             f"記錄於 {str(order_row.get('寫入時間', ''))[-5:]}" if order_row else "傳完團膳記得按「記錄」", bool(order_row)),
        card("13:10", "卡機刷卡",
             f"{counts_top.get('刷卡總數', '')} 筆" if isinstance(counts_top, dict) else "等卡機資料",
             f"環電 {counts_top.get('環電', '')}・工研院 {counts_top.get('工研院', '')}" if isinstance(counts_top, dict)
             else "13:10 自動上傳", True if isinstance(counts_top, dict) else None),
        card("13:30", "下午結算",
             f"現金 {status['cash']} 份" if status.get("cash") else "還沒記錄",
             "已寫回趨勢紀錄" if status.get("cash") else "填現金、便當盒後按「記錄」", bool(status.get("cash"))),
        card("即時", "用餐預測", f'<a href="{DASHBOARD_URL}" target="_blank">開啟戰情室 ↗</a>', "常客進園、派單量、命中率", None),
    ]
    st.markdown(f'<div class="scards">{"".join(cards)}</div>', unsafe_allow_html=True)
if status.get("error"):
    st.caption(f"⚠️ 今日進度讀取失敗：{status['error']}")
if is_friday:
    st.warning("📢 今天是星期五：下班前請到「⚙️ 名單與菜單」分頁完成下週菜單上架。")

# ==========================================
# 3. 雲端名單與菜單
# ==========================================
try:
    df_veg = load_csv_data(VEG_CSV_URL)
    normal_rice_list = df_veg[df_veg['飯量偏好'].astype(str).str.contains('正常')]['姓名'].tolist()
    no_rice_list = df_veg[df_veg['飯量偏好'].astype(str).str.contains('不要')]['姓名'].tolist()
    veg_error = None
except Exception as e:
    normal_rice_list, no_rice_list, veg_error = [], [], e

try:
    df_names = load_csv_data(NAMES_CSV_URL).fillna("")
    def names_of(kind, default_only=False):
        d = df_names[df_names['類別'].astype(str).str.strip() == kind]
        if default_only:
            d = d[d['每天預設帶入'].astype(str).str.strip() == '是']
        return [str(n).strip() for n in d['姓名'] if str(n).strip()]
    names_error = None
except Exception as e:
    names_error = e
    def names_of(kind, default_only=False):
        return []

tab_am, tab_prep, tab_pm, tab_card, tab_set = st.tabs(["☀️ 早上訂餐", "🍳 備餐打飯", "💰 下午結算", "🪪 卡片／補登", "⚙️ 名單與菜單"])

# ==========================================
# 4. 早上訂餐
# ==========================================
with tab_am:
    left, right = st.columns([5, 6], gap="large")
    with left:
        st.subheader("今天要訂多少")
        # 早上已按「記錄」：重開網頁時自動帶回當時送出的內容（臨時加量、原因、主菜、換菜、方便素）
        rec = order_row or {}
        def rec_int(k, d=0):
            try:
                return int(str(rec.get(k, "")).strip())
            except ValueError:
                return d
        rec_reason = str(rec.get("加量原因", "")).split("：", 1) if rec else [""]
        rec_sides = dict(re.match(r"^(.*?)(\d+)$", s).groups() for s in str(rec.get("換菜明細", "")).split("、")
                         if re.match(r"^(.*?)(\d+)$", s))
        rec_veg = load_veg_log(today_str) if rec else []
        if rec:
            st.success(f"已帶入早上 {str(rec.get('寫入時間', ''))[-5:]} 記錄的內容（{rec.get('總訂餐份數', '')} 份），要改直接修改再按一次記錄。")
        first_version, first_basis = first_version_base(selected_date)
        default_base = max(0, first_version - len(normal_rice_list) - len(no_rice_list)) if first_version else 0
        if rec and rec_int("實際正常數量", -1) >= 0:
            default_base = rec_int("實際正常數量") - rec_int("臨時加量")   # 以早上記錄為準，避免預測值事後變動
        if first_version:
            st.metric("📈 戰情室預測（桶餐）", f"{default_base} 份",
                      help=f"戰情室第一版 {first_version} 份（{first_basis}）− 方便素 {len(normal_rice_list) + len(no_rice_list)} 份")
        else:
            default_base = st.number_input("實際正常數量（份）", min_value=0, value=0, step=1,
                                           help="今天抓不到戰情室預測，請直接填")
        a1, a2 = st.columns([1, 2])
        with a1:
            temp_add = st.number_input("臨時加減量（份）", value=rec_int("臨時加量"), step=1, key=f"temp_add_{today_str}",
                                       help="施工廠商、訪客、活動等沒刷卡的人；人變少就填負數")
        with a2:
            reasons = ["", "進園人數比預測多", "進園人數比預測少", "施工廠商／臨時人員", "活動、外賓參訪", "部門開會訂便當（減）", "其他"]
            reason_pick = st.selectbox("原因", reasons, key=f"reason_{today_str}",
                                       index=reasons.index(rec_reason[0]) if rec_reason[0] in reasons else 0)
        reason_note = st.text_input("補充說明（選填）", placeholder="例：施工 5 人", key=f"reason_note_{today_str}",
                                    value=rec_reason[1] if len(rec_reason) > 1 else "") if temp_add else ""
        temp_reason = "：".join(s for s in (reason_pick, reason_note.strip()) if s)
        base_count = max(0, default_base + temp_add)
        if temp_add:
            st.caption(f"實際正常數量：{default_base} {'+' if temp_add > 0 else '−'} {abs(temp_add)} = **{base_count} 份**")
            if not temp_reason.strip():
                st.warning("請寫一下加量原因，之後才看得出臨時人員的規律。")

        st.markdown("**🥬 方便素**（當天沒來的按 ✕ 拿掉）")
        if veg_error:
            st.warning(f"⚠️ 方便素名單載入失敗，請手動確認。{veg_error}")
        listed = set(normal_rice_list) | set(no_rice_list)
        rv = lambda rice, in_list: [r["姓名"] for r in rec_veg if r["飯量"] == rice and ((r["姓名"] in listed) == in_list)]
        veg_normal = st.multiselect("正常飯", options=normal_rice_list, key=f"veg_n_{today_str}",
                                    default=[n for n in rv("正常飯", True)] if rec_veg else normal_rice_list)
        veg_no_rice = st.multiselect("不要白飯", options=no_rice_list, key=f"veg_x_{today_str}",
                                     default=[n for n in rv("不要白飯", True)] if rec_veg else no_rice_list)
        t1, t2 = st.columns(2)
        with t1:
            temp_veg_normal = split_names(st.text_input("➕ 臨時加訂・正常飯", placeholder="不在名單的人，多人用、隔開",
                                                        key=f"veg_tn_{today_str}", value="、".join(rv("正常飯", False))))
        with t2:
            temp_veg_no_rice = split_names(st.text_input("➕ 臨時加訂・不要白飯", placeholder="例：王小明、陳大華",
                                                         key=f"veg_tx_{today_str}", value="、".join(rv("不要白飯", False))))
        veg_listed = set(veg_normal) | set(veg_no_rice)
        veg_normal = veg_normal + [n for n in temp_veg_normal if n not in veg_listed]
        veg_no_rice = veg_no_rice + [n for n in temp_veg_no_rice if n not in veg_listed]

        st.markdown("**🍱 今日菜色加購**")
        main_dish = "查無主菜"
        main_type = ""
        extra_main_count = 0
        side_dishes = {}
        try:
            df_menu = load_csv_data(MENU_CSV_URL)
            df_risk = load_csv_data(RISK_CSV_URL)
            df_menu['橫向日期'] = pd.to_datetime(df_menu['日期']).dt.strftime('%Y-%m-%d')
            today_menu = df_menu[df_menu['橫向日期'] == today_str]
            if not today_menu.empty:
                main_dish = today_menu['主菜'].values[0]
                sides = [today_menu[f'配菜{i}'].values[0] for i in range(1, 5)]
                if '星期' in today_menu.columns:
                    display_date = f"{today_str} ({today_menu['星期'].values[0]})"
                m1, m2 = st.columns([3, 2])
                with m1:
                    main_type = st.radio(f"主菜：{main_dish}", MAIN_TYPES, horizontal=True, key=f"main_type_{today_str}",
                                         index=1 if str(rec.get("主菜類型", "")).startswith("散肉") else 0,
                                         help="散肉自助時容易多夾，通常要多加購；記下來之後可以算出散肉日平均要加多少")
                with m2:
                    extra_main_count = st.number_input("加購主菜（30 元／份）", min_value=0, step=1, key=f"main_n_{today_str}",
                                                       value=rec_int("加購主菜", 10) if rec else 10)
                st.caption("配菜（10 元）依風險係數自動建議：🔴 A 級 🟡 一般 🟢 C 級保底")
                side_cols = st.columns(2)
                for i, side in enumerate(sides):
                    suggested_amount = int(round((base_count * 0.15) / 5.0) * 5.0)
                    mark = ""
                    match_risk = df_risk[df_risk['菜色關鍵字'] == side]
                    if not match_risk.empty:
                        coef = float(match_risk['加購係數'].values[0])
                        risk_level = str(match_risk['風險等級'].values[0])
                        suggested_amount = int(round(base_count * coef / 5.0) * 5.0)
                        if "A級" in risk_level:
                            mark = "🔴"
                        elif "C級" in risk_level:
                            suggested_amount = max(10, suggested_amount)
                            mark = "🟢"
                        else:
                            mark = "🟡"
                    with side_cols[i % 2]:
                        side_dishes[side] = st.number_input(f"{mark} {side}".strip(), min_value=0, step=1, key=f"side_{i}_{today_str}",
                                                            value=int(rec_sides.get(side, 0)) if rec_sides else suggested_amount)
            else:
                st.error(f"找不到 {today_str} 的菜單，請確認日期，或到「⚙️ 名單與菜單」更新菜單試算表。")
        except Exception as e:
            st.error(f"菜單載入失敗，請檢查網路或試算表權限。{e}")

    # 配菜總額湊齊 80 元倍數（折合便當數）
    initial_side_cost = sum(side_dishes.values()) * 10 if side_dishes else 0
    if initial_side_cost % 80 != 0:
        extra_side_count = math.ceil(initial_side_cost / 80)
        diff_portions = int((extra_side_count * 80 - initial_side_cost) / 10)
        first_dish = list(side_dishes.keys())[0]
        side_dishes[first_dish] += diff_portions
        with left:
            st.info(f"💡 配菜總額湊齊 80 元倍數，已將差額 {diff_portions} 份補入「{first_dish}」。")
        extra_side_cost = extra_side_count * 80
    else:
        extra_side_count = initial_side_cost // 80
        extra_side_cost = initial_side_cost

    veg_total = len(veg_normal) + len(veg_no_rice)
    bucket_total = base_count + extra_side_count
    grand_total = bucket_total + veg_total

    side_details_str = "".join(f"\n    🥬 {k}：{v} 份" for k, v in side_dishes.items() if v > 0)

    morning_msg = f"""【 📅 {display_date} 中創園區訂餐明細 】

🎯 今日總訂餐份數：{grand_total} 份 
(桶餐 {bucket_total} 份 ＋ 方便素 {veg_total} 份)

🍱 一、 桶餐明細 (共 {bucket_total} 份)
* 實際正常數量：{base_count} 份 + ( {extra_side_count} 份換菜)
* 🥩 額外加購主菜：{extra_main_count} 份
* 👉 換菜加購內容： {side_details_str}
    (加購總計 {extra_side_cost} 元，剛好折合 {extra_side_count} 個便當)

🥬 二、 方便素便當 (共 {veg_total} 份)
* 正常飯 ({len(veg_normal)} 份)：{bullet_list(veg_normal)}
* 不要白飯 ({len(veg_no_rice)} 份)：{bullet_list(veg_no_rice)}

⚠️ 三、 廚房提醒事項
* 🔴 請務必預留「檢體一份」"""

    with right:
        st.subheader("傳給團膳群組")
        st.caption("按框框右上角的 📋 圖示複製，貼到團膳群組。")
        st.code(morning_msg, language="text")
        if has_gs_secret():
            if st.button(f"💾 記錄今天的訂餐（{grand_total} 份）", type="primary", use_container_width=True):
                try:
                    side_str = "、".join(f"{k}{v}" for k, v in side_dishes.items() if v > 0)
                    st.success("✅ " + write_order_part(today_str, [
                        today_str, first_version or "", base_count, extra_side_count, veg_total,
                        base_count + veg_total, grand_total, datetime.now(TAIPEI).strftime("%Y-%m-%d %H:%M"),
                        extra_main_count, temp_add, temp_reason.strip(), main_dish, main_type.split("（")[0], side_str,
                    ], AM_COLS, "早上訂餐"))
                    st.success("✅ " + write_veg_log(today_str, [(n, "正常飯") for n in veg_normal] +
                                                     [(n, "不要白飯") for n in veg_no_rice],
                                                     set(normal_rice_list) | set(no_rice_list)))
                    load_today_status.clear()
                    load_veg_log.clear()
                except Exception as e:
                    st.error(f"寫入失敗：{e}")
            st.caption("這個按鈕只會存紀錄，不會傳任何訊息。明細請照上面複製貼到團膳群組，貼完再按一下。下午結算和預測分析會用到，同一天再按一次會覆蓋。")
        else:
            st.warning("尚未設定試算表金鑰，無法記錄送出份數。")

# ==========================================
# 5. 備餐打飯
# ==========================================
with tab_prep:
    left, right = st.columns([5, 6], gap="large")
    with left:
        st.subheader("今天的外帶與餐盤")
        if names_error:
            st.warning(f"⚠️ 外帶餐盤名單載入失敗，請手動選擇。{names_error}")
        st.markdown(f"**🥡 方便素外帶**：{veg_total} 份（自動帶入早上訂餐的方便素名單）")
        meat_opts = names_of("葷食外帶")
        meat_takeout_names = st.multiselect("🍱 葷食外帶", options=meat_opts, default=names_of("葷食外帶", True),
                                            key=f"meat_{today_str}")
        # 現場餐盤：固定每天用餐的人自動帶入，請假時按 ✕ 拿掉
        plate_default = list(dict.fromkeys(FIXED_PLATE + names_of("餐盤", True)))
        plate_opts = list(dict.fromkeys(plate_default + names_of("餐盤")))
        plate_names = st.multiselect("🍽️ 現場餐盤", options=plate_opts, default=plate_default, key=f"plate_{today_str}")
        st.caption("請假的人按 ✕ 拿掉。要加人或改每天預設，到「⚙️ 名單與菜單」打開試算表修改。")

    veg_details_list = [f"{n}(正常飯)" for n in veg_normal] + [f"{n}(不要飯)" for n in veg_no_rice]
    prep_msg = f"""【 📅 {display_date} 備餐與打飯明細 】

大家好，今日需協助打包與裝盤的明細如下：

一、 🥡 方便素外帶 (共 {veg_total} 份)
* 名單：{bullet_list(veg_details_list)}

二、 🍱 葷食外帶 (共 {len(meat_takeout_names)} 份)
* 名單：{bullet_list(meat_takeout_names)}

三、 🍽️ 餐盤裝盛 (共 {len(plate_names)} 份)
* 名單：{bullet_list(plate_names)}

辛苦了，謝謝！"""

    with right:
        st.subheader("09:30 傳給內部備餐群組")
        st.code(prep_msg, language="text")

# ==========================================
# 6. 下午結算
# ==========================================
with tab_pm:
    left, right = st.columns([5, 6], gap="large")
    with left:
        st.subheader("填下午數字")
        sent_total = int(order_row.get("總訂餐份數") or grand_total) if order_row else grand_total
        sent_main = int(order_row.get("加購主菜") or extra_main_count) if order_row and str(order_row.get("加購主菜", "")).strip() else extra_main_count
        # 團膳只接受早上一版：總供餐數、加購主菜直接用早上送出的數字，特殊狀況才打開改
        if order_row:
            st.markdown(f"總供餐 **{sent_total}** 份　加購主菜 **{sent_main}** 份（早上已記錄）")
        else:
            st.warning("早上沒有按「記錄」，下面先用「早上訂餐」分頁目前的數字，請對照早上傳給團膳的明細確認。")
        with st.expander("✏️ 份數和早上不一樣才需要改", expanded=not order_row):
            s1, s2 = st.columns(2)
            with s1:
                total_served = st.number_input("總供餐數（份）", min_value=0, value=sent_total, step=1)
            with s2:
                main_served = st.number_input("加購主菜（份）", min_value=0, value=sent_main, step=1)
        adjust = total_served - sent_total

        counts_note = None
        default_card = default_hd = default_agl = None
        if has_gs_secret():
            counts = load_daily_counts(today_str)
            if isinstance(counts, dict):
                default_card = int(counts.get("工研院") or 0)
                default_hd = int(counts.get("環電") or 0)
                default_agl = int(counts.get("奧鋼聯") or 0)
                st.success(f"✅ 卡機統計已自動帶入（{counts.get('更新時間', '')}）")
            elif counts is None:
                st.info("⏳ 今天的卡機統計還沒產生，請手動填寫。")
            else:
                st.warning(f"⚠️ 讀取卡機統計失敗，請手動填寫。{counts}")
            if st.button("🔄 重新讀取卡機統計"):
                load_daily_counts.clear()
                st.rerun()
        # 卡機數字自動帶入時收起來，只有數字不對才需要打開改
        auto = default_hd is not None
        if auto:
            st.markdown(f"工研院 **{default_card}** 人　環電 **{default_hd}** 人　奧鋼聯 **{default_agl}** 人")
        with st.expander("✏️ 卡機數字不對才需要改", expanded=not auto):
            k1, k2, k3 = st.columns(3)
            with k1:
                card_count = st.number_input("工研院刷卡（人）", min_value=0, value=default_card, step=1, placeholder="請填寫")
            with k2:
                hd_count = st.number_input("環電（人）", min_value=0, value=default_hd, step=1, placeholder="請填寫")
            with k3:
                agl_count = st.number_input("奧鋼聯（人）", min_value=0, value=default_agl, step=1, placeholder="請填寫")
        st.markdown("**請填：**")
        c1, c2 = st.columns(2)
        with c1:
            cash_count = st.number_input("現場付現（人）", min_value=0, value=None, step=1, placeholder="請填寫")
        with c2:
            box_count = st.number_input("加購外帶便當盒（組）", min_value=0, value=None, step=1, placeholder="請填寫")

    missing = [n for n, v in [("現場付現", cash_count), ("便當盒", box_count), ("環電", hd_count)] if v is None]
    with right:
        st.subheader("13:00 傳給團膳群組")
        if missing:
            st.warning(f"⚠️ 還沒填：{'、'.join(missing)}。填好後才會顯示結算明細，避免用到錯的數字。")
        else:
            total_meal_cost = total_served * 80
            extra_main_cost = main_served * 30
            cash_deduction = cash_count * 80
            box_cost = box_count * 5
            total_cash_handover = cash_deduction + box_cost
            final_payment = total_meal_cost + extra_main_cost - cash_deduction
            afternoon_msg = f"""【 💰 {display_date} 中創園區午餐結算明細 】

一、 總供餐費用
* 總供餐數：{total_served} 份 
* 小計：{total_served} 份 × 80 元 = {total_meal_cost:,} 元

二、 額外加主菜
* 今日加主菜：{main_served} 份 
* 小計：{main_served} 份 × 30 元 = {extra_main_cost:,} 元

三、 現場付現交接明細
* 現金付費便當：{cash_count} 份 × 80 元 = {cash_deduction:,} 元
* 加購外帶便當盒：{box_count} 組 × 5 元 = {box_cost:,} 元
* ⚠️ 總交接現金：{total_cash_handover:,} 元 (已備妥，請於下午回收餐桶時一併核對，並於紙本簽名後帶走)

🎯 四、 今日最終結帳總額 (不含便當盒代收付)
➡️ 團膳請款金額：{final_payment:,} 元

🏢 五、 環電今日用餐人數
* 總計：{hd_count} 人"""
            st.code(afternoon_msg, language="text")
            # 現金份數回寫「趨勢紀錄」（戰情室的實際用餐 = 刷卡 + 現金，明天第一版會用到），結算數字記進「團膳訂單」
            if has_gs_secret():
                if st.button(f"💾 記錄今日結算（現金 {cash_count} 份）", type="primary", use_container_width=True):
                    try:
                        st.success("✅ " + write_cash_to_trend(today_str, cash_count))
                        st.success("✅ " + write_order_part(today_str, [
                            adjust, total_served, main_served, cash_count, box_count, hd_count,
                            datetime.now(TAIPEI).strftime("%Y-%m-%d %H:%M"),
                        ], PM_COLS, "下午結算"))
                        load_today_status.clear()
                    except Exception as e:
                        st.error(f"寫入失敗：{e}")
                st.caption("只會存紀錄，不會傳訊息。現金份數會變成明天以後第一版的計算基礎，請每天都要按。")

# ==========================================
# 7. 新增廠商卡片：寫進「新增卡片申請」，子健筆電每 10 分鐘處理（SQL、卡機名單、主檔），完成後 LINE 回報
# ==========================================
with tab_card:
    st.subheader("新增進駐廠商卡片")
    st.caption("送出後由筆電自動寫入資料庫並更新卡機，約 10 分鐘內完成，完成會收到 LINE。"
               "11:00～13:30 用餐時段不更新卡機，會等 13:30 後處理。")
    if not has_gs_secret():
        st.warning("尚未設定試算表金鑰，無法送出申請。")
    else:
        with st.form("card_form", clear_on_submit=True):
            f1, f2 = st.columns(2)
            with f1:
                req_emp = st.text_input("員工編號", placeholder="例：CTIC-128-375")
                req_name = st.text_input("姓名", placeholder="例：蔡起杰")
            with f2:
                req_comp = st.text_input("公司", placeholder="例：環電")
                req_card = st.text_input("卡號（10 碼）", placeholder="例：2510440043")
            submitted = st.form_submit_button("📨 送出申請", type="primary")
        if submitted:
            emp, name, comp, card = (s.strip() for s in (req_emp, req_name, req_comp, req_card))
            problems = [m for ok, m in [(emp, "員工編號"), (name, "姓名"), (comp, "公司"),
                                         (card.isdigit() and 8 <= len(card) <= 10, "卡號（要 8～10 位數字）")] if not ok]
            if problems:
                st.error(f"請確認：{'、'.join(problems)}")
            else:
                try:
                    book = open_access_log()
                    try:
                        ws = book.worksheet(CARD_SHEET)
                    except Exception:
                        ws = book.add_worksheet(title=CARD_SHEET, rows=500, cols=len(CARD_HEADERS))
                        ws.update(values=[CARD_HEADERS], range_name="A1")
                    ws.append_row([datetime.now(TAIPEI).strftime("%Y-%m-%d %H:%M"), emp.replace("_", "-"), name, comp,
                                   card, "待處理", "", ""], value_input_option="RAW")
                    st.success(f"✅ 已送出：{comp} {name}（{emp}），卡號 {card}。完成後會收到 LINE。")
                    load_card_requests.clear()
                except Exception as e:
                    st.error(f"送出失敗：{e}")
        reqs = load_card_requests()
        if isinstance(reqs, list) and reqs:
            st.markdown("**最近的申請**")
            st.dataframe(pd.DataFrame(reqs[-10:][::-1]), hide_index=True, use_container_width=True)
        elif isinstance(reqs, str):
            st.caption(f"⚠️ 讀不到申請紀錄：{reqs}")
        if st.button("🔄 重新整理狀態"):
            load_card_requests.clear()
            load_forgot_requests.clear()
            st.rerun()

    # ---- 忘記帶卡人工補登（只限廠商，例如環電；寫進本機 Vendor_Access，月結請款會算到）----
    st.divider()
    st.subheader("廠商用餐補登／取消")
    st.caption("只限園區廠商（例如環電），工研院同仁要走院部流程。可以處理前幾天的。")
    fg_mode = st.radio("要做什麼", ["📝 補登：忘記帶卡，有吃沒刷到", "🚫 取消：有刷卡，但沒吃到（例如菜沒了）"],
                       horizontal=True)
    cancel = fg_mode.startswith("🚫")
    fg_sheet, fg_word = (CANCEL_SHEET, "取消") if cancel else (FORGOT_SHEET, "補登")
    st.caption("取消後，那天那一筆刷卡會從月結請款單拿掉（筆電會另外留紀錄）。" if cancel
               else "補登後，每人那天記一筆 80 元，月底請款單會算進去。")
    if has_gs_secret():
        people = load_vendor_people()
        if isinstance(people, str):
            st.warning(f"讀不到員工名單：{people}")
        else:
            comps = sorted({p["公司"] for p in people})
            # 公司放在表單外面，換公司時姓名選單才會跟著更新
            fg_comp = st.selectbox("公司", comps, index=next((i for i, c in enumerate(comps) if "環電" in c), 0))
            with st.form("forgot_form", clear_on_submit=True):
                g2, g3 = st.columns([3, 2])
                with g2:
                    # 選項含員工編號與卡號，可以打姓名、員工編號後三碼或卡號搜尋
                    opts = {f"{p['姓名']}｜{p['員工編號']}｜卡號 {p['卡號']}": p for p in people if p["公司"] == fg_comp}
                    fg_people = st.multiselect("姓名（可打姓名、員工編號後三碼或卡號搜尋，可選多人）", list(opts))
                with g3:
                    fg_date = st.date_input("用餐日期", value=current_time.date(), max_value=current_time.date())
                fg_note = st.text_input("備註（選填）", placeholder="例：菜沒了，沒吃到" if cancel else "例：忘記帶卡，現場登記")
                fg_submit = st.form_submit_button(f"📨 送出{fg_word}", type="primary")
            if fg_submit:
                if not fg_people:
                    st.error(f"請選擇要{fg_word}的人。")
                else:
                    try:
                        ws = forgot_ws(fg_sheet)
                        now_s = datetime.now(TAIPEI).strftime("%Y-%m-%d %H:%M")
                        rows = []
                        for o in fg_people:
                            p = opts[o]
                            rows.append([now_s, f"{fg_date:%Y-%m-%d}", p["員工編號"], p["姓名"], fg_comp, "待處理", "", fg_note.strip()])
                        ws.append_rows(rows, value_input_option="RAW")
                        st.success(f"✅ 已送出 {fg_date:%m/%d} {fg_word} {len(rows)} 位，約 10 分鐘內處理，完成會收到 LINE。")
                        load_forgot_requests.clear()
                    except Exception as e:
                        st.error(f"送出失敗：{e}")
            fr = load_forgot_requests(fg_sheet)
            if isinstance(fr, list) and fr:
                st.markdown(f"**最近的{fg_word}**")
                st.dataframe(pd.DataFrame(fr[-10:][::-1]), hide_index=True, use_container_width=True)

    # ---- 姓名更正：廠商同仁建檔打錯字，連過去用餐紀錄、卡機名單一起改 ----
    st.divider()
    st.subheader("姓名更正")
    st.caption("輸入員工編號、後三碼或卡號找人，確認目前登記的姓名；打錯就填正確姓名送出。"
               "筆電會一併更正資料庫、過去所有用餐紀錄（日報、月報）、卡機名單與員工主檔。只限園區廠商。")
    if has_gs_secret():
        people_all = load_vendor_people()
        if isinstance(people_all, list):
            nm_opts = {f"{p['姓名']}｜{p['員工編號']}｜{p['公司']}｜卡號 {p['卡號']}": p for p in people_all}
            nm_pick = st.selectbox("找人（打員工編號、後三碼、卡號或姓名）", [""] + list(nm_opts))
            if nm_pick:
                p = nm_opts[nm_pick]
                i1, i2, i3 = st.columns(3)
                i1.metric("目前登記姓名", p["姓名"])
                i2.metric("員工編號", p["員工編號"])
                i3.metric("卡號", p["卡號"] or "—")
                with st.form("rename_form", clear_on_submit=True):
                    new_name = st.text_input("正確姓名", placeholder="例：陳靜萍")
                    nm_submit = st.form_submit_button("✏️ 送出更正", type="primary")
                if nm_submit:
                    nn = new_name.strip()
                    if not nn or nn == p["姓名"]:
                        st.error("請輸入和目前不同的正確姓名。")
                    else:
                        try:
                            forgot_ws(RENAME_SHEET).append_row(
                                [datetime.now(TAIPEI).strftime("%Y-%m-%d %H:%M"), "", p["員工編號"], nn, p["公司"],
                                 "待處理", "", f"原姓名：{p['姓名']}"], value_input_option="RAW")
                            st.success(f"✅ 已送出：{p['姓名']} → {nn}，約 10 分鐘內更正，完成會收到 LINE。")
                            load_forgot_requests.clear()
                        except Exception as e:
                            st.error(f"送出失敗：{e}")
            rr = load_forgot_requests(RENAME_SHEET)
            if isinstance(rr, list) and rr:
                st.markdown("**最近的更正**")
                st.dataframe(pd.DataFrame(rr[-10:][::-1]).drop(columns=["用餐日期"], errors="ignore"),
                             hide_index=True, use_container_width=True)

# ==========================================
# 8. 名單與菜單
# ==========================================
with tab_set:
    st.subheader("名單維護")
    st.markdown(f"""
- **方便素**、**葷食外帶／餐盤**名單都在 [👉 中創園區每週菜單試算表]({EDIT_SHEET_URL})，改完約 10 分鐘內生效。
  - 方便素：「姓名、飯量偏好（正常飯／不要白飯）」
  - 外帶餐盤名單：「姓名、類別（葷食外帶／餐盤）、每天預設帶入（是／否）」
""")
    if has_gs_secret():
        with st.expander("🥬 素食訂購統計（近 30 天）"):
            try:
                v = pd.DataFrame(open_access_log().worksheet(VEG_LOG_SHEET).get_all_records())
                v = v[pd.to_datetime(v["日期"]) >= pd.Timestamp(selected_date) - pd.Timedelta(days=30)]
                st.markdown(f"每天平均 **{v.groupby('日期').size().mean():.1f}** 份，其中臨時加訂 "
                            f"**{(v['名單內'] != '是').sum()}** 人次")
                stat = v.groupby(["姓名", "名單內"]).size().reset_index(name="訂購天數").sort_values("訂購天數", ascending=False)
                st.dataframe(stat, hide_index=True, use_container_width=True)
                st.caption("臨時加訂次數多的人，可以加進上面的方便素名單，每天就會自動帶入。")
            except Exception as e:
                st.caption(f"還沒有統計資料（{type(e).__name__}）")
    if st.button("🔄 立即重新讀取名單與菜單"):
        load_csv_data.clear()
        st.rerun()
    st.subheader("📌 每週五：下週菜單同步與上架")
    st.markdown(f"""
- [ ] **1. 更新雲端試算表：** [👉 中創園區雲端試算表後台]({EDIT_SHEET_URL}) 填入下週菜單。
- [ ] **2. 上架工研食堂：** [👉 工研食堂管理後台]({ITRI_CANTEEN_URL}) 上架下週主菜、配菜。
- [ ] **3. 更新中創生活網：** [👉 中創生活編輯頁面]({CTIC_LIFE_URL}) 發布下週菜單公告。
""")
