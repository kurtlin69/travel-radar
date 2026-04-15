import time
import pandas as pd
import urllib.parse
from datetime import datetime
import io
import os
import re
import streamlit as st
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup

# ================= 1. 自動路徑與配置區 =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CHROME_PROFILE_PATH = os.path.join(BASE_DIR, "Selenium_Chrome_Profile")

if not os.path.exists(CHROME_PROFILE_PATH):
    os.makedirs(CHROME_PROFILE_PATH)

HOTEL_KEYWORDS = ["優惠", "團購", "轉讓", "住宿券", "下殺", "促銷", "專案", "特價", "折扣", "補助", "早鳥", "快閃"]

# ================= 2. 核心功能函數 =================

def generate_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='統整報表')
        worksheet = writer.sheets['統整報表']
        worksheet.column_dimensions['A'].width = 12  
        worksheet.column_dimensions['B'].width = 18  
        worksheet.column_dimensions['C'].width = 18  
        worksheet.column_dimensions['D'].width = 80  
        worksheet.column_dimensions['E'].width = 60  
    return output.getvalue()

def get_driver():
    options = Options()
    options.add_argument(f"user-data-dir={CHROME_PROFILE_PATH}")
    options.add_argument("--window-position=-32000,-32000") 
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36")
    
    driver = webdriver.Chrome(options=options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
      "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    driver.implicitly_wait(10)
    return driver

# ================= 3. 各平台爬蟲模組 =================

def scrape_abic(driver, hotel_keyword):
    items = []
    try:
        driver.get("https://www.abic.com.tw/event")
        time.sleep(4) 
        soup = BeautifulSoup(driver.page_source, "html.parser")
        found_links = set()
        for link in soup.find_all("a"):
            href = link.get("href")
            if not href or href.startswith('javascript'): continue 
            combined_text = link.get_text(strip=True) + " " + (link.get("title") or "")
            if hotel_keyword in combined_text:
                if href.startswith('/'): href = "https://www.abic.com.tw" + href
                if href not in found_links:
                    found_links.add(href)
                    items.append({"抓取日期": datetime.now().strftime("%Y-%m-%d"), "目標": hotel_keyword, "平台": "愛貝客", "資訊": combined_text[:100], "連結": href})
    except: pass
    return items

def scrape_gomaji(driver, keyword):
    items = []
    try:
        encoded_keyword = urllib.parse.quote(keyword)
        driver.get(f"https://www.gomaji.com/search?keyword={encoded_keyword}")
        time.sleep(4)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        found_links = set()
        for link in soup.find_all("a"):
            href = link.get("href")
            if not href or "gomaji.com" not in href: continue
            text = link.get_text(strip=True)
            if keyword in text and href not in found_links:
                found_links.add(href)
                items.append({"抓取日期": datetime.now().strftime("%Y-%m-%d"), "目標": keyword, "平台": "GOMAJI", "資訊": text[:100], "連結": href})
    except: pass
    return items

def scrape_momo(driver, hotel_keyword):
    items = []
    try:
        encoded_keyword = urllib.parse.quote(hotel_keyword + " 住宿券")
        driver.get(f"https://www.momoshop.com.tw/search/searchShop.jsp?keyword={encoded_keyword}")
        time.sleep(4) 
        soup = BeautifulSoup(driver.page_source, "html.parser")
        for link in soup.find_all("a"):
            href = link.get("href")
            if not href or "goodsUrl" not in href: continue
            text = link.get_text(strip=True)
            if hotel_keyword in text:
                if href.startswith('/'): href = "https://www.momoshop.com.tw" + href
                items.append({"抓取日期": datetime.now().strftime("%Y-%m-%d"), "目標": hotel_keyword, "平台": "MOMO", "資訊": text[:100], "連結": href})
    except: pass
    return items

def scrape_eztravel(driver, hotel_keyword):
    items = []
    try:
        encoded_keyword = urllib.parse.quote(hotel_keyword)
        driver.get(f"https://www.eztravel.com.tw/search/?q={encoded_keyword}")
        time.sleep(5) 
        soup = BeautifulSoup(driver.page_source, "html.parser")
        for link in soup.find_all("a"):
            href = link.get("href")
            text = link.get_text(strip=True)
            if hotel_keyword in text:
                items.append({"抓取日期": datetime.now().strftime("%Y-%m-%d"), "目標": hotel_keyword, "平台": "易遊網", "資訊": text[:100], "連結": href if href.startswith("http") else f"https://www.eztravel.com.tw{href}"})
    except: pass
    return items

def scrape_fb_groups(driver, keyword):
    posts = []
    try:
        driver.get("https://www.facebook.com/")
        time.sleep(3) 
        encoded_keyword = urllib.parse.quote(keyword)
        search_url = f"https://www.facebook.com/search/posts/?q={encoded_keyword}"
        driver.get(search_url)
        time.sleep(8) 
        articles = driver.find_elements(By.XPATH, "//div[@role='article']")
        for article in articles:
            content = article.text.strip()
            if keyword in content and any(word in content for word in HOTEL_KEYWORDS):
                posts.append({"抓取日期": datetime.now().strftime("%Y-%m-%d"), "目標": keyword, "平台": "FB 社群", "資訊": content[:150].replace('\n', ' '), "連結": search_url})
    except: pass
    return posts

def scrape_google_maps_food(driver, keyword):
    items = []
    try:
        search_query = keyword if any(x in keyword for x in ["美食", "小吃", "餐廳"]) else f"{keyword} 美食"
        encoded_keyword = urllib.parse.quote(search_query)
        driver.get(f"https://www.google.com/maps/search/{encoded_keyword}")
        
        time.sleep(8) 
        soup = BeautifulSoup(driver.page_source, "html.parser")
        found_names = set()
        
        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            if "/maps/place/" in href:
                name = a_tag.get("aria-label")
                if not name or name in found_names: continue
                
                rating = ""
                parent_text = a_tag.parent.get_text() if a_tag.parent else ""
                rating_match = re.search(r'([0-4]\.\d|5\.0)\s*顆星', parent_text)
                if rating_match: rating = f" (⭐ {rating_match.group(1)} 顆星)"
                
                found_names.add(name)
                items.append({"抓取日期": datetime.now().strftime("%Y-%m-%d"), "目標": search_query, "平台": "Google Maps", "資訊": f"在地推薦：【{name}】{rating}", "連結": href})
                if len(items) >= 15: break
    except Exception as e: st.error(f"地圖抓取失敗: {e}")
    return items

# ================= 全新：深度評價分析引擎 =================
def deep_analyze_target(driver, target):
    """搜尋目標評價並進行基礎的情緒分析"""
    try:
        # 鎖定搜尋真實評價與避雷關鍵字
        search_query = f"{target} 評價 OR 缺點 OR 踩雷 OR 推薦 (PTT OR Dcard OR 食記)"
        encoded_query = urllib.parse.quote(search_query)
        driver.get(f"https://www.google.com/search?q={encoded_query}")
        time.sleep(4)
        
        soup = BeautifulSoup(driver.page_source, "html.parser")
        snippets = []
        
        # 抓取 Google 搜尋結果的文字摘要 (Snippet)
        for block in soup.find_all("div", class_="VwiC3b"):
            text = block.get_text(strip=True)
            if len(text) > 20: snippets.append(text)
            
        if not snippets:
            return "❌ 無法蒐集到足夠的網路評價資料，建議擴大關鍵字範圍。"
            
        combined_text = " ".join(snippets)
        
        # 情緒關鍵字字典
        pos_words = ["好吃", "推薦", "讚", "CP值", "必吃", "美味", "乾淨", "舒服", "親切", "回訪", "入味", "神"]
        neg_words = ["雷", "難吃", "態度差", "貴", "不推", "失望", "髒", "普通", "過譽", "沒味道", "死鹹", "蟑螂"]
        
        pos_count = sum(combined_text.count(w) for w in pos_words)
        neg_count = sum(combined_text.count(w) for w in neg_words)
        
        total = pos_count + neg_count
        
        # 演算法判定
        if total == 0:
            verdict = "💬 **討論度平淡**：網路討論較少情緒性字眼，建議親自體驗或參考 Google 實體星級。"
        elif pos_count > neg_count * 2.5:
            verdict = "🔥 **強烈推薦！** 網路評價一面倒的好，絕對值得一試，趕快記下來！"
        elif pos_count > neg_count:
            verdict = "👍 **值得一吃！** 大多數人給予正面評價，踩雷機率低。"
        elif pos_count == neg_count:
            verdict = "🤔 **評價兩極！** 有人很愛有人覺得還好，建議確認自己是否能接受某些特定缺點。"
        else:
            verdict = "⚠️ **高機率踩雷！** 負面避雷關鍵字偏多，除非有執念，否則建議考慮其他選項。"

        # 組裝報告
        report = f"### 🕵️ 【{target}】 深度調查報告\n\n"
        report += f"{verdict}\n\n"
        report += f"---\n**📊 關鍵字數據分析：**\n"
        report += f"- 🟢 正向討論熱度：**{pos_count}** 分\n"
        report += f"- 🔴 負面避雷熱度：**{neg_count}** 分\n\n"
        report += f"**📝 擷取網友真實評價碎片：**\n"
        
        for i, snip in enumerate(snippets[:3]):
            clean_snip = snip.replace("...", "").strip()
            report += f"> 「*{clean_snip[:80]}...*」\n\n"
            
        return report
        
    except Exception as e:
        return f"分析過程發生異常: {e}"

# ================= 4. Streamlit 主畫面 =================

st.set_page_config(page_title="國內旅遊情報搜查雷達", layout="wide", page_icon="🧳")

if 'search_results' not in st.session_state:
    st.session_state.search_results = None
if 'report_type' not in st.session_state:
    st.session_state.report_type = "旅遊情報"
if 'analysis_report' not in st.session_state:
    st.session_state.analysis_report = ""

# --- 側邊欄：深度評論探測器 ---
with st.sidebar:
    st.header("⚙️ 系統控制")
    if st.button("🚪 關閉系統", type="primary", use_container_width=True):
        st.success("伺服器終止中...")
        time.sleep(2)
        os._exit(0) 
        
    st.divider()
    
    st.header("🕵️ 深度評價探測器")
    st.markdown("怕踩雷？輸入店名，系統幫你掃描全網評價並給出最終判定！")
    
    target_to_analyze = st.text_input("輸入想調查的店名/飯店", placeholder="例如: 龍爺牛肉麵")
    
    if st.button("🔍 啟動深度分析", use_container_width=True):
        if not target_to_analyze:
            st.warning("請先輸入名稱！")
        else:
            with st.spinner(f"正在全網探勘 {target_to_analyze} 的真實評價..."):
                driver = None
                try:
                    driver = get_driver()
                    result = deep_analyze_target(driver, target_to_analyze)
                    st.session_state.analysis_report = result
                finally:
                    if driver: driver.quit()
    
    # 顯示分析報告
    if st.session_state.analysis_report:
        st.info(st.session_state.analysis_report)
        st.download_button(
            label="📥 下載此分析報告",
            data=st.session_state.analysis_report.encode('utf-8'),
            file_name=f"{target_to_analyze}_避雷分析報告.txt",
            mime="text/plain",
            use_container_width=True
        )

# --- 主畫面區塊 ---
st.title("🧳 國內旅遊情報搜查雷達")
st.markdown(f"運行目錄：`{BASE_DIR}`")

tab1, tab2 = st.tabs(["🏨 住宿優惠搜尋", "🗺️ 在地美食地圖探索"])

with tab1:
    st.markdown("### 🏨 尋找理想飯店的晚鳥或專案優惠")
    hotel_input = st.text_input("🔍 輸入飯店名稱 (用逗號分隔)", placeholder="例如: 捷絲旅, 蘭城晶英", key="hotel_input")
    
    if st.button("🚀 開始搜尋住宿", type="primary"):
        target_list = [h.strip() for h in hotel_input.split(",") if h.strip()]
        if not target_list: st.warning("請輸入名稱")
        else:
            all_data = []
            driver = None
            my_bar = st.progress(0, text="準備中...")
            try:
                driver = get_driver()
                for hotel in target_list:
                    for platform, func in [("愛貝客", scrape_abic), ("GOMAJI", scrape_gomaji), ("MOMO", scrape_momo), ("易遊網", scrape_eztravel), ("FB社群", scrape_fb_groups)]:
                        my_bar.progress(0.5, text=f"搜尋 【{hotel}】 - {platform}...")
                        all_data.extend(func(driver, hotel))
                my_bar.progress(1.0, text="整理中...")
            finally:
                if driver: driver.quit()
            
            if all_data:
                st.session_state.search_results = pd.DataFrame(all_data).drop_duplicates(subset=['資訊'])
                st.session_state.report_type = "🏨住宿優惠" 
                st.success(f"找到 {len(st.session_state.search_results)} 筆優惠")
            else: st.info("無結果")

with tab2:
    st.markdown("### 🗺️ Google Maps 真實評價探索")
    food_input = st.text_input("🔍 輸入地區或美食", placeholder="例如: 頭份, 宜蘭 小吃", key="food_input")
    
    if st.button("📍 啟動地圖探測", type="primary"):
        target_list = [f.strip() for f in food_input.split(",") if f.strip()]
        if not target_list: st.warning("請輸入地名")
        else:
            all_data = []
            driver = None
            try:
                driver = get_driver()
                for food in target_list:
                    all_data.extend(scrape_google_maps_food(driver, food))
            finally:
                if driver: driver.quit()
            
            if all_data:
                st.session_state.search_results = pd.DataFrame(all_data).drop_duplicates(subset=['資訊'])
                st.session_state.report_type = "🍜美食地標" 
                st.success(f"鎖定 {len(st.session_state.search_results)} 個地標")
            else: st.info("無結果")

if st.session_state.search_results is not None:
    st.divider()
    st.subheader(f"📊 {st.session_state.report_type}搜查結果")
    st.dataframe(st.session_state.search_results, use_container_width=True)
    
    excel_data = generate_excel_bytes(st.session_state.search_results)
    dynamic_filename = f"{st.session_state.report_type}報表_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    
    st.download_button(
        label=f"📥 下載 {st.session_state.report_type} Excel 報表", 
        data=excel_data, 
        file_name=dynamic_filename, 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )