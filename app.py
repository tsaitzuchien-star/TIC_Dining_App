import streamlit as st
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
import math

TAIPEI = ZoneInfo("Asia/Taipei")

# ==========================================
# 0. 網頁基本設定 & 快取功能 (🌟 解決 409 衝突的核心)
# ==========================================
st.set_page_config(page_title="中創園區訂餐 App", page_icon="🍱", layout="wide")

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
ORDER_SHEET = "團膳訂單"        # 早上送出的份數，戰情室 09:20 用來算追加／減量
# 趨勢紀錄的發佈 CSV（不需要金鑰），用來算戰情室的 08:30 第一版
TREND_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSv8X_NJDPjSVJTEkFjdBTSJKjDnZ18eyf6yBdTBq2eVT8TB6_wB-UDUUVUoF5eUJrLSnDsqpu7OhbS/pub?gid=1825976142&single=true&output=csv"
DASHBOARD_URL = "https://tsaitzuchien-star.github.io/CTIC_Dining_Prediction/"
FALLBACK_BASE = {0: 132, 1: 128, 2: 119, 3: 126, 4: 108}   # 趨勢紀錄資料不足時的星期基準量
BASE_LOOKBACK = 8

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

def write_order(row):
    """把早上送出的份數寫進「團膳訂單」，同一天再按一次就覆蓋"""
    ws = open_access_log().worksheet(ORDER_SHEET)
    dates = ws.col_values(1)
    if row[0] in dates:
        r = dates.index(row[0]) + 1
        ws.update(values=[row], range_name=f"A{r}:I{r}", value_input_option="USER_ENTERED")
        return f"已更新 {row[0]} 的送出份數"
    ws.append_row(row, value_input_option="USER_ENTERED")
    return f"已記錄 {row[0]} 的送出份數"

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
col_title, col_date = st.columns([3, 1])
with col_title:
    st.title("🍱 中創園區訂餐 App")
    st.markdown(f"今日份數由 [📊 用餐預測戰情室]({DASHBOARD_URL}) 自動帶入，這裡負責排明細、傳 LINE。")
with col_date:
    selected_date = st.date_input("📅 日期", value=current_time.date())
today_str = selected_date.strftime("%Y-%m-%d")
backup_week = weekdays_ch[selected_date.weekday()]
display_date = f"{today_str} ({backup_week})"
is_friday = selected_date.weekday() == 4

status = load_today_status(today_str) if has_gs_secret() else {}
order_row = status.get("order")
if has_gs_secret():
    p1, p2, p3 = st.columns(3)
    with p1:
        if order_row:
            st.success(f"08:30 ✅ 已記錄送出 {order_row.get('總訂餐份數', '')} 份（{str(order_row.get('寫入時間', ''))[-5:]}）")
        else:
            st.warning("08:30 ⬜ 早上明細還沒記錄")
    with p2:
        st.info(f"09:20 追加／減量短訊請到 [戰情室]({DASHBOARD_URL}) 複製")
    with p3:
        if status.get("cash"):
            st.success(f"13:00 ✅ 現金 {status['cash']} 份已寫回")
        else:
            st.warning("13:00 ⬜ 下午現金還沒寫回")
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

tab_am, tab_prep, tab_pm, tab_set = st.tabs(["☀️ 早上訂餐", "🍳 備餐打飯", "💰 下午結算", "⚙️ 名單與菜單"])

# ==========================================
# 4. 早上訂餐
# ==========================================
with tab_am:
    left, right = st.columns([5, 6], gap="large")
    with left:
        st.subheader("今天要訂多少")
        first_version, first_basis = first_version_base(selected_date)
        default_base = max(0, first_version - len(normal_rice_list) - len(no_rice_list)) if first_version else 0
        base_count = st.number_input("實際正常數量（份）", min_value=0, value=default_base, step=1)
        if first_version:
            st.caption(f"📈 戰情室第一版 {first_version} 份（{first_basis}）− 方便素 {len(normal_rice_list) + len(no_rice_list)} 份 = {default_base}。有活動可直接改。")

        st.markdown("**🥬 方便素**（當天沒來的按 ✕ 拿掉）")
        if veg_error:
            st.warning(f"⚠️ 方便素名單載入失敗，請手動確認。{veg_error}")
        veg_normal = st.multiselect("正常飯", options=normal_rice_list, default=normal_rice_list)
        veg_no_rice = st.multiselect("不要白飯", options=no_rice_list, default=no_rice_list)

        st.markdown("**🍱 今日菜色加購**")
        main_dish = "查無主菜"
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
                extra_main_count = st.number_input(f"加購主菜：{main_dish}（30 元）", min_value=0, value=10, step=1)
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
                        side_dishes[side] = st.number_input(f"{mark} {side}".strip(), min_value=0, value=suggested_amount, step=1)
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
* 🔴 請務必預留「檢體一份」
* 📢 將視今天入園人數於 09:20 前，回報是否追加餐點與今日最終數量。追加部分放在「補菜桶」即可。"""

    with right:
        st.subheader("傳給兩個團膳群組")
        st.markdown("1. 👥 **工研院 強心臟組（家常在）**　2. 👥 **FY114 家常在x工研院**")
        st.caption("按框框右上角的 📋 圖示複製，分別貼到兩個群組。")
        st.code(morning_msg, language="text")
        if has_gs_secret():
            if st.button(f"💾 已傳給團膳，記錄 {grand_total} 份", type="primary", use_container_width=True):
                try:
                    st.success("✅ " + write_order([today_str, first_version or "", base_count, extra_side_count, veg_total,
                                                    base_count + veg_total, grand_total,
                                                    datetime.now(TAIPEI).strftime("%Y-%m-%d %H:%M"), extra_main_count]))
                    load_today_status.clear()
                except Exception as e:
                    st.error(f"寫入失敗：{e}")
            st.caption("09:20 的追加／減量短訊請到戰情室複製，不用重貼這份明細。同一天再按一次會覆蓋。")
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
        st.markdown(f"**🥡 方便素外帶**：{veg_total} 份（依早上訂餐分頁）")
        meat_opts = names_of("葷食外帶")
        meat_takeout_names = st.multiselect("🍱 葷食外帶", options=meat_opts, default=names_of("葷食外帶", True))
        plate_opts = names_of("餐盤")
        plate_names = st.multiselect("🍽️ 現場餐盤", options=plate_opts, default=names_of("餐盤", True))
        st.caption("要加人或改每天預設，到「⚙️ 名單與菜單」打開試算表修改。")

    veg_details_list = [f"{n}(正常飯)" for n in veg_normal] + [f"{n}(不要飯)" for n in veg_no_rice]
    prep_msg = f"""【 📅 {display_date} 備餐與打飯明細 】

慧萍、小容、玉玲 妳們好，今日需協助打包與裝盤的明細如下：

一、 🥡 方便素外帶 (共 {veg_total} 份)
* 名單：{bullet_list(veg_details_list)}

二、 🍱 葷食外帶 (共 {len(meat_takeout_names)} 份)
* 名單：{bullet_list(meat_takeout_names)}

三、 🍽️ 餐盤裝盛 (共 {len(plate_names)} 份)
* 名單：{bullet_list(plate_names)}

辛苦了，謝謝！"""

    with right:
        st.subheader("09:30 傳給內部備餐群組")
        st.markdown("👥 **玉玲 Ling, 子健《秉澔&秉宸》, 黃慧萍, 欣柔, 小容**")
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
        if order_row:
            st.caption(f"早上已記錄：總訂餐 {sent_total} 份、加購主菜 {sent_main} 份。")
        else:
            st.caption("早上沒有記錄，先用「早上訂餐」分頁目前的數字，請確認。")
        adjust = st.number_input("09:20 追加（+）／減量（−）份數", value=0, step=1,
                                 help="照戰情室 09:20 短訊填；沒有追加就填 0")
        total_served = st.number_input("總供餐數（份）", min_value=0, value=max(0, sent_total + adjust), step=1)
        main_served = st.number_input("加購主菜（份）", min_value=0, value=sent_main, step=1)

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
        c1, c2 = st.columns(2)
        with c1:
            card_count = st.number_input("工研院刷卡（人）", min_value=0, value=default_card, step=1, placeholder="請填寫")
            agl_count = st.number_input("奧鋼聯（人）", min_value=0, value=default_agl, step=1, placeholder="請填寫")
            cash_count = st.number_input("現場付現（人）", min_value=0, value=None, step=1, placeholder="請填寫")
        with c2:
            hd_count = st.number_input("環電（人）", min_value=0, value=default_hd, step=1, placeholder="請填寫")
            box_count = st.number_input("加購外帶便當盒（組）", min_value=0, value=None, step=1, placeholder="請填寫")

    missing = [n for n, v in [("現場付現", cash_count), ("便當盒", box_count), ("環電", hd_count)] if v is None]
    with right:
        st.subheader("13:00 傳給團膳群組")
        st.markdown("👥 **工研院 強心臟組（家常在）**")
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
            # 現金份數回寫「趨勢紀錄」，戰情室的實際用餐 = 刷卡 + 現金，明天第一版會用到
            if has_gs_secret():
                if st.button(f"💾 把現金 {cash_count} 份寫入趨勢紀錄", type="primary", use_container_width=True):
                    try:
                        st.success("✅ " + write_cash_to_trend(today_str, cash_count))
                        load_today_status.clear()
                    except Exception as e:
                        st.error(f"寫入失敗：{e}")
                st.caption("現金份數會變成明天以後第一版的計算基礎，請每天都要寫回。")

# ==========================================
# 7. 名單與菜單
# ==========================================
with tab_set:
    st.subheader("名單維護")
    st.markdown(f"""
- **方便素**、**葷食外帶／餐盤**名單都在 [👉 中創園區每週菜單試算表]({EDIT_SHEET_URL})，改完約 10 分鐘內生效。
  - 方便素：「姓名、飯量偏好（正常飯／不要白飯）」
  - 外帶餐盤名單：「姓名、類別（葷食外帶／餐盤）、每天預設帶入（是／否）」
""")
    if st.button("🔄 立即重新讀取名單與菜單"):
        load_csv_data.clear()
        st.rerun()
    st.subheader("📌 每週五：下週菜單同步與上架")
    st.markdown(f"""
- [ ] **1. 更新雲端試算表：** [👉 中創園區雲端試算表後台]({EDIT_SHEET_URL}) 填入下週菜單。
- [ ] **2. 上架工研食堂：** [👉 工研食堂管理後台]({ITRI_CANTEEN_URL}) 上架下週主菜、配菜。
- [ ] **3. 更新中創生活網：** [👉 中創生活編輯頁面]({CTIC_LIFE_URL}) 發布下週菜單公告。
""")
