import streamlit as st
import pandas as pd
from datetime import datetime
import math
import pyodbc  # SQL Server 連線套件

# ==========================================
# 0. 網頁基本設定
# ==========================================
st.set_page_config(page_title="中創園區用餐預測戰情室", page_icon="🍱", layout="centered")

st.title("🍱 中創園區用餐預測戰情室")
st.markdown("負責每日與「家常在」團膳業者的自動化訂餐與結算系統")

# ==========================================
# 1. 抓取系統時間與設定雲端資料庫網址
# ==========================================
current_time = datetime.now()
today_str = current_time.strftime("%Y-%m-%d")

MENU_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vT9QdhFOdM2cp7FI1qu4VNRvwOF6mHDJZ7OP0iYTu2shMiF5PrZI3lUzyP436KyBV3uv49akqBytF47/pub?output=csv"
VEG_CSV_URL = "https://docs.google.com/spreadsheets/d/1dGsbEe6aCJo0gexj5Xo2gdTmQ_oA6E4VNIdmEHZDGZM/export?format=csv&gid=1496853361"
RISK_CSV_URL = "https://docs.google.com/spreadsheets/d/1dGsbEe6aCJo0gexj5Xo2gdTmQ_oA6E4VNIdmEHZDGZM/export?format=csv&gid=2090477701"

EDIT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1dGsbEe6aCJo0gexj5Xo2gdTmQ_oA6E4VNIdmEHZDGZM/edit"
ITRI_CANTEEN_URL = "https://restorg.itri.org.tw/index.aspx"
CTIC_LIFE_URL = "https://sites.google.com/d/1QxzEC936SJmN9fFzheP6cbl43eveS2FY/p/1VcccyLG4yToXJgO56nnF93Q-xuLg233w/edit?pli=1"

# ==========================================
# 🌟 每週五固定待辦事項 
# ==========================================
is_friday = (current_time.weekday() == 4)

with st.expander("📌 每週五必做任務：下週菜單同步與上架檢核表", expanded=is_friday):
    if is_friday:
        st.warning("📢 ⚠️ 提醒：今天是星期五！請務必在今天下班前完成以下菜單更新流程。")
    else:
        st.info("💡 備忘提示：每週五需固定執行的菜單上架流程如下。")
        
    st.markdown(f"""
    - [ ] **1. 更新雲端試算表：** 收到菜單後，點擊進入 [👉 中創園區雲端試算表後台]({EDIT_SHEET_URL}) 填入下週菜單。
    - [ ] **2. 上架工研食堂：** 點擊進入 [👉 工研食堂管理後台]({ITRI_CANTEEN_URL}) 將下週的各項主菜、配菜公告上架。
    - [ ] **3. 更新中創生活網：** 點擊進入 [👉 中創生活編輯頁面]({CTIC_LIFE_URL}) 發布下週菜單公告，開放給全園區同仁知悉。
    """)

# ==========================================
# 2. 側邊欄：參數設定與雲端資料載入
# ==========================================
with st.sidebar:
    st.header("⚙️ 今日參數設定")
    
    selected_date = st.date_input("📅 選擇日期", value=current_time)
    today_str = selected_date.strftime("%Y-%m-%d")
    
    weekdays_ch = ["一", "二", "三", "四", "五", "六", "日"]
    backup_week = weekdays_ch[selected_date.weekday()]
    display_date = f"{today_str} ({backup_week})"
    
    base_count = st.number_input("1. 實際正常數量 (份)", min_value=0, value=140, step=1)
    
    st.subheader("2. 方便素名單 (打包外帶)")
    try:
        df_veg = pd.read_csv(VEG_CSV_URL)
        normal_rice_list = df_veg[df_veg['飯量偏好'].astype(str).str.contains('正常')]['姓名'].tolist()
        no_rice_list = df_veg[df_veg['飯量偏好'].astype(str).str.contains('不要')]['姓名'].tolist()
        
        veg_normal = st.multiselect("正常飯", options=normal_rice_list, default=normal_rice_list)
        veg_no_rice = st.multiselect("不要白飯", options=no_rice_list, default=no_rice_list)
    except Exception as e:
        st.warning("⚠️ 吃素名單載入失敗，使用預設值。")
        veg_normal = st.multiselect("正常飯", ["蔡ＯＯ", "林ＯＯ"])
        veg_no_rice = st.multiselect("不要白飯", ["陳ＯＯ", "王ＯＯ"])

    st.subheader("3. 🥡 葷食外帶名單")
    meat_takeout_names = st.multiselect("預定葷食外帶人員", 
                                        options=["國聯鄭釗文", "國聯吳光豪", "歐米巴", "臨時新增B"], 
                                        default=["國聯鄭釗文", "歐米巴"])

    st.subheader("4. 🍽️ 現場餐盤名單")
    plate_names = st.multiselect("預定餐盤人員", 
                                 options=["高健維", "臨時新增C", "臨時新增D"], 
                                 default=["高健維"])

    st.subheader("5. 🍱 今日菜色自動加購")
    
    main_dish = "查無主菜"
    extra_main_count = 10
    side_dishes = {} 
    
    try:
        df_menu = pd.read_csv(MENU_CSV_URL)
        df_risk = pd.read_csv(RISK_CSV_URL)
        
        df_menu['橫向日期'] = pd.to_datetime(df_menu['日期']).dt.strftime('%Y-%m-%d')
        today_menu = df_menu[df_menu['橫向日期'] == today_str]
        
        if not today_menu.empty:
            main_dish = today_menu['主菜'].values[0]
            sides = [today_menu['配菜1'].values[0], today_menu['配菜2'].values[0], 
                     today_menu['配菜3'].values[0], today_menu['配菜4'].values[0]]
            
            if '星期' in today_menu.columns:
                sheet_week = today_menu['星期'].values[0]
                display_date = f"{today_str} ({sheet_week})"
            
            st.success("✅ 雲端菜單與星期自動載入成功")
            
            st.markdown(f"**🥩 主菜 (30元):** {main_dish}")
            extra_main_count = st.number_input(f"加購 {main_dish} 份數", min_value=0, value=10, step=1)
            
            st.markdown("**🥬 配菜 (10元):**")
            
            for side in sides:
                default_coef = 0.15 
                suggested_amount = int(round((base_count * default_coef) / 5.0) * 5.0)
                warning_msg = ""
                
                match_risk = df_risk[df_risk['菜色關鍵字'] == side]
                if not match_risk.empty:
                    coef = float(match_risk['加購係數'].values[0])
                    risk_level = str(match_risk['風險等級'].values[0])
                    raw_calc = base_count * coef
                    suggested_amount = int(round(raw_calc / 5.0) * 5.0)
                    
                    if "A級" in risk_level:
                        warning_msg = f"🔴 (動態係數 {int(coef*100)}%)"
                    elif "C級" in risk_level:
                        suggested_amount = max(10, suggested_amount) 
                        warning_msg = f"🟢 (保底機制)"
                    else:
                        warning_msg = f"🟡 (動態係數 {int(coef*100)}%)"
                
                side_dishes[side] = st.number_input(f"{side} {warning_msg}", min_value=0, value=suggested_amount, step=1)
                
        else:
            st.error(f"找不到 {today_str} 的菜單！請確認日期。")
            
    except Exception as e:
         st.error(f"連線失敗，請檢查網址或權限。")

    st.subheader("6. 下午結算專用 (財務記帳)")
    cash_count = st.number_input("現場付現 (人數)", min_value=0, value=15, step=1)
    
    # 🌟 新增：便當盒加購欄位
    box_count = st.number_input("加購外帶便當盒 (組)", min_value=0, value=4, step=1)
    
    card_count = st.number_input("工研院刷卡 (人數)", min_value=0, value=46, step=1)
    hd_count = st.number_input("環電 (人數)", min_value=0, value=36, step=1)
    agl_count = st.number_input("奧鋼聯 (人數)", min_value=0, value=0, step=1)

# ==========================================
# 3. 背景邏輯計算與字串排版
# ==========================================
initial_side_cost = sum(side_dishes.values()) * 10 if side_dishes else 0

if initial_side_cost % 80 != 0:
    extra_side_count = math.ceil(initial_side_cost / 80)
    adjusted_side_cost = extra_side_count * 80
    diff_cost = adjusted_side_cost - initial_side_cost 
    diff_portions = int(diff_cost / 10) 
    
    if side_dishes:
        first_dish = list(side_dishes.keys())[0]
        side_dishes[first_dish] += diff_portions
        st.sidebar.info(f"💡 財務防呆：總金額湊齊 80 倍數，系統已自動將差額 ({diff_portions} 份) 補入「{first_dish}」中。")
    
    extra_side_cost = adjusted_side_cost
else:
    extra_side_count = initial_side_cost // 80
    extra_side_cost = initial_side_cost  

veg_total = len(veg_normal) + len(veg_no_rice)
bucket_total = base_count + extra_side_count
grand_total = bucket_total + veg_total

# 名字條列式排版處理 (對內備餐用)
veg_details_list = []
for name in veg_normal:
    veg_details_list.append(f"{name}(正常飯)")
for name in veg_no_rice:
    veg_details_list.append(f"{name}(不要飯)")
veg_details_str = "\n    - " + "\n    - ".join(veg_details_list) if veg_details_list else "無"

meat_details_str = "\n    - " + "\n    - ".join(meat_takeout_names) if meat_takeout_names else "無"
plate_details_str = "\n    - " + "\n    - ".join(plate_names) if plate_names else "無"

# 方便素分類條列式排版 (對外訂餐用)
veg_normal_str = "\n    - " + "\n    - ".join(veg_normal) if veg_normal else "無"
veg_no_rice_str = "\n    - " + "\n    - ".join(veg_no_rice) if veg_no_rice else "無"

# 🌟 財務邏輯更新：分開計算「公司請款」與「現場現金交接」
total_meal_cost = grand_total * 80
extra_main_cost = extra_main_count * 30

cash_deduction = cash_count * 80  # 便當收現
box_cost = box_count * 5          # 額外紙盒收現
total_cash_handover = cash_deduction + box_cost  # 下午要交給團膳大哥的實體現金總額

card_deduction = card_count * 80
hd_deduction = hd_count * 80
agl_deduction = agl_count * 80

daily_difference = total_meal_cost - cash_deduction - card_deduction - hd_deduction - agl_deduction
final_payment = total_meal_cost + extra_main_cost - cash_deduction  # 公司月底匯款金額不變，只扣掉便當現金

# ==========================================
# 4. 早上 08:30 - 09:00 訂餐發送區
# ==========================================
st.header("☀️ 早上 08:30 - 09:00 訂餐發送區")

st.markdown("""⚠️ **請「分別發送」至以下兩個 LINE 群組：**
1. 👥 **工研院 強心臟組（家常在）**
2. 👥 **FY114 家常在x工研院**""")

side_details_str = ""
for k, v in side_dishes.items():
    if v > 0:
        side_details_str += f"\n    🥬 {k}：{v} 份"

morning_msg = f"""【 📅 {display_date} 中創園區訂餐明細 】

🎯 今日總訂餐份數：{grand_total} 份 
(桶餐 {bucket_total} 份 ＋ 方便素 {veg_total} 份)

🍱 一、 桶餐明細 (共 {bucket_total} 份)
* 實際正常數量：{base_count} 份 + ( {extra_side_count} 份換菜)
* 🥩 額外加購主菜：{extra_main_count} 份
* 👉 換菜加購內容： {side_details_str}
    (加購總計 {extra_side_cost} 元，剛好折合 {extra_side_count} 個便當)

🥬 二、 方便素便當 (共 {veg_total} 份)
* 正常飯 ({len(veg_normal)} 份)：{veg_normal_str}
* 不要白飯 ({len(veg_no_rice)} 份)：{veg_no_rice_str}

⚠️ 三、 廚房提醒事項
* 🔴 請務必預留「檢體一份」
* 📢 將視今天入園人數於 09:20 前，回報是否追加餐點與今日最終數量。追加部分放在「補菜桶」即可。"""

st.code(morning_msg, language="text")

st.divider()

# ==========================================
# 5. 早上 09:30 - 10:00 備餐打飯發送區
# ==========================================
st.header("🍳 早上 09:30 - 10:00 備餐打飯發送區")

st.markdown("""⚠️ **請發送至以下 LINE 群組：**
1. 👥 **玉玲 Ling, 子健《秉澔&秉宸》, 黃慧萍, 華苓, 欣柔, 小容**""")

prep_msg = f"""【 📅 {display_date} 備餐與打飯明細 】

慧萍、小容、玉玲 妳們好，今日需協助打包與裝盤的明細如下：

一、 🥡 方便素外帶 (共 {veg_total} 份)
* 名單：{veg_details_str}

二、 🍱 葷食外帶 (共 {len(meat_takeout_names)} 份)
* 名單：{meat_details_str}

三、 🍽️ 餐盤裝盛 (共 {len(plate_names)} 份)
* 名單：{plate_details_str}

辛苦了，謝謝！"""

st.code(prep_msg, language="text")

st.divider()

# ==========================================
# 6. 下午 13:00 - 14:00 結算發送區
# ==========================================
st.header("💰 下午 13:00 - 14:00 結算發送區")

st.markdown("""⚠️ **請發送至以下 LINE 群組：**
1. 👥 **工研院 強心臟組（家常在）**""")

# 🌟 更新結算訊息，清楚交代便當盒的現金去向
afternoon_msg = f"""【 💰 {display_date} 中創園區午餐結算明細 】

一、 總供餐費用
* 總供餐數：{grand_total} 份 
* 小計：{grand_total} 份 × 80 元 = {total_meal_cost:,} 元

二、 額外加主菜
* 今日加主菜：{extra_main_count} 份 (固定主菜)
* 小計：{extra_main_count} 份 × 30 元 = {extra_main_cost:,} 元

三、 現場付現交接明細
* 現金付費便當：{cash_count} 份 × 80 元 = {cash_deduction:,} 元
* 加購外帶便當盒：{box_count} 組 × 5 元 = {box_cost:,} 元
* ⚠️ 總交接現金：{total_cash_handover:,} 元 (已備妥，請於下午回收廚餘時一併核對，並於紙本簽名後帶走)

🎯 四、 今日最終結帳總額 (不含便當盒代收付)
➡️ 團膳請款金額：{final_payment:,} 元"""

st.code(afternoon_msg, language="text")

st.divider()

# ==========================================
# 7. 雲端帳務資料庫 (防呆查帳專區)
# ==========================================
st.header("📊 雲端帳務資料庫 (防呆複製區)")
st.markdown("""👉 **月底對帳救星！** 複製下方文字，到「中創園區_每日帳務資料庫」試算表的空白儲存格按下 `Ctrl + V`，即可自動分格！""")
audit_row = f"{today_str}\t{backup_week}\t{grand_total}\t{cash_deduction}\t{daily_difference}\t{card_count}\t{agl_count}\t{hd_count}"
st.code(audit_row, language="text")

st.divider()

# ==========================================
# 8. 月底結算報表匯出區 (專屬環電、奧鋼聯)
# ==========================================
st.header("🗄️ 月底結算自動撈取區 (環電與奧鋼聯)")
st.markdown("直接從 SQL Server `Vendor_Access` 資料表撈取指定區間的加總數據，拒絕人工計算錯誤！工研院資料請直接至工研食堂後台匯出。")

with st.expander("點擊展開：產生 SQL 結算報表", expanded=False):
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("結算起始日期", value=datetime(2026, 4, 1))
    with col2:
        end_date = st.date_input("結算結束日期", value=datetime(2026, 4, 30))

    if st.button("🚀 一鍵撈取 SQL 資料"):
        try:
            SERVER = r'14-0A00035-93\SQLEXPRESS'
            DATABASE = 'Dining' 
            
            DB_DATE_COL = "AccessDate" 
            DB_VENDOR_COL = "DepName"  
            
            conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};Trusted_Connection=yes;'
            conn = pyodbc.connect(conn_str)
            
            sql_query = f"""
            SELECT 
                CAST([{DB_DATE_COL}] AS DATE) AS 結帳日期,
                [{DB_VENDOR_COL}] AS 公司名稱, 
                COUNT(*) AS 訂餐總份數
            FROM 
                Vendor_Access
            WHERE 
                CAST([{DB_DATE_COL}] AS DATE) >= '{start_date}' 
                AND CAST([{DB_DATE_COL}] AS DATE) <= '{end_date}'
                AND [{DB_VENDOR_COL}] IN ('環電', '環電股份有限公司', '奧鋼聯')
            GROUP BY 
                CAST([{DB_DATE_COL}] AS DATE), 
                [{DB_VENDOR_COL}]
            ORDER BY 
                結帳日期 ASC;
            """
            
            df_sql = pd.read_sql(sql_query, conn)
            conn.close()
            
            if df_sql.empty:
                st.warning("這段期間內沒有環電或奧鋼聯的訂餐紀錄喔！")
            else:
                st.success("✅ 資料庫撈取成功！以下為精準結算數據：")
                st.dataframe(df_sql, use_container_width=True)
                
                csv = df_sql.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 下載對帳表 (CSV檔)",
                    data=csv,
                    file_name=f"環電_奧鋼聯_結算表_{start_date}至{end_date}.csv",
                    mime="text/csv",
                )
        except Exception as e:
            st.error(f"連線失敗！請檢查錯誤訊息：{e}")
