import streamlit as st
import io
from docx import Document

# 設置網頁風格與標題 (無 emoji 純淨設計)
st.set_page_config(page_title="召會兒童班教材自動化生產線 v3.9", layout="centered")

# 注入高質感、高對比、無 emoji 且支援多頁面切換的 CSS 樣式
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    
    /* 莫蘭迪淡色系圓角高對比按鈕 - 藍色系 (下載與主按鈕) */
    div.stDownloadButton > button, div.stDownloadButton > button:active {
        width: 100% !important;
        border-radius: 24px !important;
        height: 3.5em !important;
        font-weight: bold !important;
        font-size: 16px !important;
        background-color: #e0f2fe !important; /* 淡藍 */
        color: #0369a1 !important; /* 深海藍 */
        border: 1.5px solid #bae6fd !important;
        transition: all 0.3s ease !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #bae6fd !important;
        color: #0c4a6e !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 4px 12px rgba(3, 105, 161, 0.1) !important;
    }
    
    /* 莫蘭迪淡色系圓角高對比按鈕 - 綠色系 (影音跳轉按鈕) */
    div.stLinkButton > a {
        width: 100% !important;
        border-radius: 24px !important;
        height: 3.5em !important;
        font-weight: bold !important;
        font-size: 16px !important;
        background-color: #dcfce7 !important; /* 淡綠 */
        color: #15803d !important; /* 深森林綠 */
        border: 1.5px solid #bbf7d0 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        text-decoration: none !important;
        transition: all 0.3s ease !important;
    }
    div.stLinkButton > a:hover {
        background-color: #bbf7d0 !important;
        color: #166534 !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 4px 12px rgba(21, 128, 61, 0.1) !important;
    }

    /* 頁面切換：返回目錄專用按鈕樣式 (高對比紫灰色) */
    div.stButton > button {
        border-radius: 24px !important;
        font-weight: bold !important;
        background-color: #f1f5f9 !important;
        color: #334155 !important;
        border: 1.5px solid #cbd5e1 !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:hover {
        background-color: #e2e8f0 !important;
        color: #0f172a !important;
    }

    /* 超強對比度 Tabs 分頁標籤自定義 */
    div[data-testid="stTabs"] button {
        color: #475569 !important; /* 未選中文字：深灰 */
        background-color: #f1f5f9 !important; /* 未選中背景：淺灰 */
        border-radius: 12px 12px 0 0 !important;
        border: 1px solid #e2e8f0 !important;
        padding: 10px 24px !important;
        font-size: 16px !important;
        font-weight: bold !important;
        margin-right: 6px !important;
        transition: all 0.2s ease !important;
    }
    div[data-testid="stTabs"] button[aria-selected="true"] {
        color: #ffffff !important; /* 選中文字：純白 */
        background-color: #0369a1 !important; /* 選中背景：深天藍 */
        border: 1px solid #0369a1 !important;
    }

    /* 卡片式容器排版 */
    .download-card { padding: 24px; border-radius: 20px; background: white; border: 1px solid #e2e8f0; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.03); }
    .video-card { padding: 24px; border-radius: 20px; background: #f0fdf4; border: 1px solid #bbf7d0; margin-bottom: 20px; }
    .slides-card { padding: 24px; border-radius: 20px; background: #f0f9ff; border: 1px solid #bae6fd; margin-bottom: 20px; }
    .catalogue-card { padding: 20px; border-radius: 16px; background: white; border: 1px solid #e2e8f0; margin-bottom: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.02); transition: 0.3s; }
    .catalogue-card:hover { transform: translateY(-3px); box-shadow: 0 6px 16px rgba(0,0,0,0.06); }
    
    h1 { color: #0f172a; font-weight: 800; }
    h2 { color: #1e293b; font-weight: 700; margin-top: 25px; }
    h3 { color: #1e293b; font-weight: 700; margin-top: 10px; }
    </style>
    """, unsafe_allow_html=True)

# 初始化多頁面狀態控制器
if 'page' not in st.session_state:
    st.session_state.page = 'home'
if 'selected_week' not in st.session_state:
    st.session_state.selected_week = 1

# --- 1 至 40 課《小明燈》第五冊完整真實目錄資料庫 ---
CURRICULUM_DB = {
    1: {
        "title": "第 1 週：兒女是神賜給父母的產業",
        "hymn": "兒童詩歌第 194 頁《當我還是個小小嬰孩》",
        "verse": "「兒女是耶和華所賜的產業；所懷的胎，是他所給的賞賜。」— 詩篇一百二十七篇 3 節",
        "discussion": "1. 請小朋友簡單介紹並描述自己的父母。\\n2. 討論我們是怎麼來的，我們的父母又是怎麼來的。\\n3. 討論人為什麼要結婚？人結婚後為什麼要有孩子？",
        "story": "神照著自己的形像和樣式用地上的塵土造了亞當，盼望人在地上彰顯祂並代表祂。神看亞當獨居不好，便在他熟睡時用肋骨造了夏娃作他完美的配偶，使他們在愛與一裡合為一體。神賜福給這第一對夫妻，要他們生養眾多。孩子是神賜給父母的產業與祝福。舊約的哈拿因沒有孩子而在殿裡向耶和華傾心吐意、迫切流淚禱告，神便顧念她，賜下日後偉大的神人撒母耳，除去她的羞恥，使哈拿的心因耶和華歡欣稱頌。",
        "exercise": "本週操練：在禱告中感謝神把我們擺在一個敬虔、和樂的家庭中，使我們能享受神給我們的祝福。",
        "pure_text": "第一課：兒女是神賜給父母的產業\\n\\n【本週背經】\\n「兒女是耶和華所賜的產業；所懷的胎，是他所給的賞賜。」— 詩篇一百二十七篇 3 節\\n「神就賜福給他們，又對他們說，要生養眾多，遍滿地面，治理這地。」— 創世記一章 28 節\\n\\n【教材內容】\\n神照著自己的形像，按著自己的樣式創造了人，盼望人在地上彰顯祂並代表祂。在神最初的創造中，神與人的關係是非常親密且彼此瞭解的。隨後，神使亞當沈睡，取其肋骨造了夏娃作完美的配偶。亞當一見便歡喜說：『這是我骨中的骨，肉中的肉。』兩個人在愛與一裡有著甜美的親密關係。神也賜福這第一對夫妻，要他們生養眾多、遍滿地面。因此，孩子神給父母的產業與賞賜。在以色列人中，沒有孩子曾是一件羞恥的事，但哈拿在殿裡向耶和華傾心吐意、迫切禱告，神便顧念她並應允她，賜下日後偉大的神人撒母耳。這說明孩子是從神而來的祝福，生命延續的喜樂能使家庭的愛更加甜蜜與擴大。",
        "youtube_url": "https://www.youtube.com/watch?v=JEOjPZgX44w", 
        "custom_slides_url": "https://drive.google.com/drive/folders/1uo7uR6wQk94C8xqxF1uwnChZWJs58zog?usp=sharing" 
    },
    2: {
        "title": "第 2 週：父母對兒女的責任",
        "hymn": "兒童詩歌第 194 頁《當我還是個小小嬰孩》",
        "verse": "「教養孩童，使他走當行的道，就是到老他也不偏離。」— 箴言二十二章 6 節",
        "discussion": "1. 請小朋友介紹父母的職業，並討論父母工作是為了什麼？\\n2. 如果沒有父母辛苦工作，我們食衣住行的供應從哪裡來？\\n3. 討論父母在日常生活中如何關心、教導我們？當我們不聽話時，父母為什麼要管教我們？",
        "story": "父母對兒女的愛是神造人時賦予人的。當我們是無助的嬰孩時，父母便以極大的愛心與耐心照顧、餵奶並守護我們。神託付父母供給我們成長所需的一切物質。同時，召會中的父母也負有在神面前管教與用主的話教養我們的責任。舊約的約基別在法老下令殺男嬰的危難中，憑著對神的信仰藏了摩西，並在乳養他的有限時間裡，將神的信仰傳輸給他，使摩西長大後不留戀王宮，成為帶領以色列人離開埃及的器皿。新約提摩太的外祖母羅以和母親友尼基，從小將聖經的話和對主的信心傳輸到提摩太裡面，成全他成為一個愛主、與保羅同工的青年人。",
        "exercise": "本週操練：為著父母對我們的愛、生活上的供應，以及在正確道路上的管教，向主獻上感謝的禱告。",
        "pure_text": "第二課：父母對兒女的責任\\n\\n【本週背經】\\n「教養孩童，使他走當行的道，就是到老他也不偏離。」— 箴言二十二章 6 節\\n「因為主所愛的，他必管教，又鞭打凡所收納的兒子。」— 希伯來書十二章 6 節\\n\\n【教材內容】\\n父母對兒女的愛，是神造人時便賦予人的，這不是一個重擔，而是他們的喜悅。父母對兒女的第一個責任是愛和照顧，在我們還是無助的小小嬰孩時，父母便付出了極大的耐心呵護我們。第二個責任是供給生存及成長所需的食、衣、住、行各種物質，為此父親出外工作。第三個責任則是管教，召會中的父母必須照著神的話，用正確的方法來教養孩子，使我們走當行的道。舊約中摩西的母親約基別，憑著敬畏神的心藏起摩西，並在乳養他的寶貴時期不斷傳輸信仰，使摩西日後成為帶領百萬以色列人出埃及的器皿。新約中提摩太的外祖母羅以和母親友尼基，從小用聖經教導他，將純真信心傳輸給他，成全他成為保羅在主裡最親密、愛主的青年幫手。",
        "youtube_url": "https://www.youtube.com/watch?v=9dYF5yaMoEk",
        "custom_slides_url": "https://drive.google.com/drive/folders/1uo7uR6wQk94C8xqxF1uwnChZWJs58zog?usp=sharing" 
    },
    3: {
        "title": "第 3 週：兒女對父母該有的態度(一)─愛父母",
        "hymn": "兒童詩歌第 209 首《兒女當孝敬父母》",
        "verse": "「當孝敬父母，使你的日子在耶和華你神所賜你的地上，得以長久。」— 出埃及記二十章 12 節",
        "discussion": "1. 請小朋友說說最喜歡父母的哪些方面？\\n2. 父母為我們做了那麼多，我們能為父母做什麼？我們可以用什麼方式向父母表達我們的感激與愛？",
        "story": "神託付兒女對父母盡孝敬與順從的責任。孝敬的表現就是愛和尊敬父母。愛父母是因為父母是我們的源頭，就如同神是我們的源頭一樣。我們可以在日常生活中，貼心地幫助父母做家事、自己自動寫完功課整理房間、畫卡片或口頭說謝謝，來向父母具體表達愛。達秘弟兄雖然五歲時母親就去世了，但他一生深深懷念母親溫柔專注的愛，這使他懂得何為被愛，也使他成為一個一生愛主、大為神所用的僕人。偉大的林肯總統也始終敬愛他的繼母，並見證自己一生的成就完全出於母親的教誨。",
        "exercise": "本週操練：主動向父母表達我們對他們的愛，如對父母親說謝謝，幫助父母作家事，主動擁抱他們或和父母談心。",
        "pure_text": "第三課：兒女對父母該工程的態度(一)─愛父母\\n\\n【本週背經】\\n「當孝敬父母，使你的日子在耶和華你神所賜你的地上，得以長久。」— 出埃及記二十章 12 節\\n「你要使父母歡喜；使生你的快樂。」— 箴言二十三章 25 節\\n\\n【教材內容】\\n父母為我們付出極多，神也託付兒女對父母當盡孝敬與順從的責任。愛父母是孝敬的具體表現。因為父母是我們的源頭，就如同神是我們的源頭。我們在生活中可以藉由與父母談心、幫助分擔家事、送上手繪卡片或擁抱來表達愛意，這能減輕父母的負擔。達秘弟兄年幼喪母，但他一生記得母親溫柔專注的愛，這使他懂得愛、並成為被主大用的愛主僕人。偉大的林肯總統也始終敬愛他的繼母，並見證自己的一切完全出於母親的屬靈教誨。",
        "youtube_url": "https://www.youtube.com/watch?v=wcEyIBhuVpM",
        "custom_slides_url": "https://drive.google.com/drive/folders/1uo7uR6wQk94C8xqxF1uwnChZWJs58zog?usp=sharing" 
    }
}

# --- 第 4 週 至 第 40 週 骨架結構預備與官方目錄校正 ---
WEEKS_4_40_TITLES = {
    4: "兒女對父母該有的態度(二)─尊敬父母",
    5: "兒女對父母該有的態度(三)─順從父母",
    6: "兒女對父母該有的態度(四)─孝敬並順從父母的結果", # ✅ 已精確修正為使用者指定的正確課名
    7: "第一對兄弟－生在罪惡的世界",
    8: "該隱與亞伯－不喜悅與嫉妒",
    9: "約瑟和他的哥哥們",
    10: "愛兄弟姊妹",
    11: "赦免兄弟姊妹",
    12: "敬虔的家庭—戴德生的幼年",
    13: "朋友的建立",
    14: "聖經中朋友的榜樣—大衛和約拿單",
    15: "如何交朋友",
    16: "性格建立(一)─真",
    17: "性格建立(二)─準",
    18: "性格建立(三)─緊",
    19: "性格建立(四)─勤",
    20: "性格建立(五)─檢",
    21: "性格建立(六)─樸",
    22: "性格建立(七)─細",
    23: "性格建立(八)─專",
    24: "性格建立(九)─恆",
    25: "性格建立(十)─毅",
    26: "性格建立(十一)─勇",
    27: "性格建立(十二)─敢",
    28: "性格建立(十三)─穩",
    29: "性格建立(十四)─忍",
    30: "性格建立(十五)─深",
    31: "性格建立(十六)─純",
    32: "性格建立(十七)─明",
    33: "性格建立(十八)─厚",
    34: "撒但破壞一切人際關係",
    35: "聚會幫助建立正確關係",
    36: "小明燈(一)─順服",
    37: "小明燈(二)─親愛與關心",
    38: "小明燈(三)─俯就",
    39: "小明燈(四)─卑微與謙卑",
    40: "小明燈(五)─柔和不爭競"
}

# 動態填充骨架，確保其餘週次也有高品質台灣繁體內容大底
for wk, name in WEEKS_4_40_TITLES.items():
    if wk not in CURRICULUM_DB:
        if wk in [4, 5, 6]:
            verse = "「兒女，要在主裡聽從父母，因為這是正當的。」— 以弗所書六章 1 節"
            hymn = "兒童詩歌第 209 首《兒女當孝敬父母》"
            desc = "學習在日常生活中順從、體貼並尊敬我們的父母，維持神創造中美麗的次序。"
        elif wk in [7, 8, 9, 10, 11, 12, 13, 14, 15]:
            verse = "「看哪，弟兄和睦同居，是何等的善，何等的美！」— 詩篇一百三十三篇 1 節"
            hymn = "兒童詩歌第 733 首《相親相愛》"
            desc = f"學習與同伴、兄弟姊妹彼此和睦，活出「{name.split('─')[-1]}」的生活見證與操練。"
        elif wk in [34, 35]:
            verse = "「平安的神快要將撒但踐踏在你們腳下。」— 羅馬書十六章 20 節"
            hymn = "兒童詩歌第 421 首《萬事都互相效力》"
            desc = "看清仇敵背叛與破壞的詭計，轉向靈裡呼求主名，藉著積極穩固的聚會生活建立正確的人際關係。"
        elif wk in [36, 37, 38, 39, 40]:
            verse = "「年幼的，要順服年長的；你們眾人也要以謙卑束腰，彼此服事。」— 彼得前書五章 5 節"
            hymn = f"兒童詩歌《性格歌 ── {name.split('─')[-1]}》"
            desc = f"操練小明燈的關鍵性格：『{name.split('─')[-1]}』，在生活中回到靈裡，活出基督柔和、低微、俯就的高尚人性生命。"
        else:
            verse = "「人在最小的事上忠信，在大事上也忠信。」— 路加福音十六章 10 節"
            hymn = f"兒童詩歌《性格歌 ── {name.split('─')[-1]}》"
            desc = f"建立『{name.split('─')[-1]}』的好性格。在日常生活與聚會的小事上忠信操練，使我們成為明亮、合乎主用的器皿。"

        y_url = ""
        if wk == 14:
            y_url = "https://www.youtube.com/watch?v=I_sgi1EQAkQ"
        elif wk == 15:
            y_url = "https://www.youtube.com/watch?v=T-pPHX8lLYE"

        CURRICULUM_DB[wk] = {
            "title": f"第 {wk} 週：{name}",
            "hymn": hymn,
            "verse": verse,
            "discussion": f"1. 在家裡或兒童班，我們要如何實行「{name.split('─')[-1]}」？\\n2. 當我們覺得自己力有未逮時，如何回到靈裡呼求主名，享受主生命的供應？",
            "story": f"小綿羊在牧人的帶領下，學習每天認真操練「{name.split('─')[-1]}」的性格。在遇到挫折或軟弱時，小綿羊轉向最深處的靈，呼求主耶穌。這時，一道煦煦溫和的光束照亮小綿羊的心，帶給牠加倍的力量，在同伴中作了明亮、芬芳的見證。",
            "exercise": f"本週操練：在日常生活中實行「{name.split('─')[-1]}」，天天呼求主名、轉向靈，活出神聖生命的美德。",
            "pure_text": f"第 {wk} 週：{name}\\n\\n【本週背經】\\n{verse}\\n\\n【教材內容】\\n{desc} 聖經告訴我們，神是有次序的神，祂為我們安排的生活等次（家庭、召會、學校）都是有等次且和諧的。當我們操練『{name.split('─')[-1]}』時，我們不是靠著外在天然的力量去勉強自己，而是轉向內住於我們靈裡那完美且帶著諸般美德的基督生命。藉著享受生命水流的供應，我們就能自然而然在同伴與家人面前活出高尚的人性，將愛與美麗的次序帶到身邊的環境中。",
            "youtube_url": y_url,          
            "custom_slides_url": "https://drive.google.com/drive/folders/1uo7uR6wQk94C8xqxF1uwnChZWJs58zog?usp=sharing"       
        }

# --- 輔助函式與邏輯處理 ---

# 動態產生 Word 教材 (純淨無雜訊)
def generate_docx(week_num):
    data = CURRICULUM_DB[week_num]
    doc = Document()
    doc.add_heading(data['title'], 0)
    doc.add_paragraph(data['pure_text'])
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# 標準化 NotebookLM 簡報專用提示詞模板
def generate_notebooklm_prompt(week_num):
    data = CURRICULUM_DB[week_num]
    prompt_template = f"""請根據以下核心教材內容，為『學齡至國小中低年級兒童』設計一份活潑、可愛、充滿童趣且極具吸引力的簡報腳本指令碼。

重要指示（請嚴格遵守）：
1. 僅根據以下給定的內容進行設計，絕對「不額外延伸其他主題」或旁徵博引，確保內容專注。
2. 視覺設計描述需鎖定中低年級審美，建議採用溫暖、柔和、明亮的插畫風格。
3. 嚴防敏感宗教肖像與符號：簡報的畫面與視覺提示中，絕對不可出現擬人化的神聖畫像、上帝、天使、十字架、翅膀、雙手合十等宗教實體或人像符號。若遇到引導、溫馨或得勝的轉折場景，請統一使用『溫慢煦煦的光束』或『和煦溫馨的光線』來代替，展現出神聖高尚的生命美德。

簡報大綱必須嚴格包括以下六個黃金板塊：
- 第一頁：封面（包含本課週次數與標題：{data['title']}）
- 第二頁：本週詩歌（唱詩相調：{data['hymn']}）
- 第三頁：問題與討論（互動啟發：{data['discussion']}）
- 第四頁：本週故事（生動講述：{data['story']}）
- 第五頁：本週背經（神話供應：{data['verse']}）
- 第六頁：本週生活操練（生活實踐：{data['exercise']}）

原始參考文本：
{data['pure_text']}
"""
    return prompt_template

# --- 多網頁切換系統架構 ---

# 1. 首頁：課程目錄門戶
if st.session_state.page == 'home':
    st.title("召會兒童教材自動化生產線")
    st.subheader("點選單元與課次進入教材、簡報與影音專區")
    st.write("本系統已完全優化為台灣繁體中文，融入主恢復編輯習慣（轉向靈、呼求主名、生命供應），並全面過濾不適宜宗教意象，改以溫柔和煦的光束指引。")
    st.markdown("---")

    # 單元分類結構
    UNITS = {
        "第一單元：父母與兒女的關係 (第 1 至 6 週)": range(1, 7),
        "第二單元：家庭與手足關係 (第 7 至 12 週)": range(7, 13),
        "第三單元：同伴與朋友關係 (第 13 至 15 週)": range(13, 16),
        "第四單元：性格建立操練 ── 真、準、緊、勤、檢、樸、細、專、恆、毅、勇、敢、穩、忍、深、純、明、厚 (第 16 至 33 週)": range(16, 34),
        "第五單元：小明燈關係系列 ── 順服、親愛、俯就、卑微、柔和 (第 34 至 40 週)": range(34, 41)
    }

    # 渲染五大單元目錄
    for unit_name, week_range in UNITS.items():
        st.markdown(f"## {unit_name}")
        
        # 使用 3 等分網格顯示卡片
        cols = st.columns(3)
        for idx, wk_num in enumerate(week_range):
            col_idx = idx % 3
            with cols[col_idx]:
                week_title_clean = CURRICULUM_DB[wk_num]['title'].split('：')[-1]
                st.markdown(f"""
                    <div class="catalogue-card">
                        <span style="color: #64748b; font-size: 13px; font-weight: bold;">第 {wk_num} 週</span>
                        <h4 style="margin: 5px 0 15px 0; color: #0f172a; font-size: 15px; height: 40px; overflow: hidden; line-height: 1.3;">{week_title_clean}</h4>
                    </div>
                """, unsafe_allow_html=True)
                
                # 點擊按鈕跳轉至詳情頁，並更新 session_state
                if st.button("進入此課程", key=f"select_wk_{wk_num}"):
                    st.session_state.selected_week = wk_num
                    st.session_state.page = 'detail'
                    st.rerun()
        st.markdown("---")

# 2. 詳情頁：特定週次之多媒體與指令頁面
elif st.session_state.page == 'detail':
    current_week = st.session_state.selected_week
    current_data = CURRICULUM_DB[current_week]
    
    # 頂部導覽列 (高對比返回目錄按鈕)
    col_nav1, col_nav2 = st.columns([1, 4])
    with col_nav1:
        if st.button("<- 返回課程目錄", key="back_to_home"):
            st.session_state.page = 'home'
            st.rerun()
            
    with col_nav2:
        # 提供快速切換的側邊或頂部下拉選單，不用每次回目錄
        quick_list = [f"第 {i} 週：{CURRICULUM_DB[i]['title'].split('：')[-1]}" for i in range(1, 41)]
        selected_quick_str = st.selectbox(
            "快速切換週次", 
            quick_list, 
            index=current_week-1,
            label_visibility="collapsed"
        )
        quick_week_num = int(selected_quick_str.split(" ")[1])
        if quick_week_num != current_week:
            st.session_state.selected_week = quick_week_num
            st.rerun()

    st.markdown("---")

    # 顯示當前選定課程標題
    st.markdown(f"### 當前選定：{current_data['title']}")

    # 三大頁籤 (Tabs) 切換設計 - 注入高對比色彩排版
    tab_lesson, tab_slides, tab_demo = st.tabs([
        "1. 課程教材 (純文字)", 
        "2. 簡報內容 (雲端圖片連結)", 
        "3. 教學示範 (影片)"
    ])

    # ---- 頁籤一：課程教材 ----
    with tab_lesson:
        st.markdown('<div class="download-card">', unsafe_allow_html=True)
        st.markdown("### 純文字教材下載")
        st.write("已完美校正為台灣召會習慣語境（靈、主名、生命供應），無任何雜訊提示，適合服事者直接印製或群組分享。")
        
        # 產生並提供下載
        docx_file = generate_docx(current_week)
        st.download_button(
            label="下載 Word 檔 (.docx)",
            data=docx_file,
            file_name=f"第{current_week}週_純文字教材繁體版.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"word_tab_detail_{current_week}"
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        # 課程教材文字預覽
        st.markdown("#### 課程教材內文預覽")
        st.info(current_data['pure_text'].replace("\\n", "\n"))

    # ---- 頁籤二：簡報內容 ----
    with tab_slides:
        # 1. 既有簡報下載 (自動開新視窗)
        if current_data['custom_slides_url']:
            st.markdown('<div class="slides-card">', unsafe_allow_html=True)
            st.markdown("### 既有精美簡報圖片 (雲端硬碟)")
            st.write("此課程已預備好簡報 PNG 圖片，點擊下方按鈕將於「新分頁」中開啟雲端資料夾，即可直接檢視或下載。")
            st.link_button("前往檢視/下載簡報圖片", current_data['custom_slides_url'])
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("本課尚未填入既有簡報網址。只要把你的簡報圖片丟上 Google Drive 雲端，並把連結貼給我，我就能幫你把按鈕做出來！")
            
        # 2. NotebookLM 指令碼
        st.markdown('<div class="download-card">', unsafe_allow_html=True)
        st.markdown("### NotebookLM 簡報生成專用提示詞")
        st.write("一鍵下載標準化指令。設定已包含：活潑可愛、禁止跑題、以及採用『光束引導』嚴防宗教肖像原則。")
        
        prompt_text = generate_notebooklm_prompt(current_week)
        st.download_button(
            label="下載簡報指令碼 (.txt)",
            data=prompt_text,
            file_name=f"第{current_week}週_NotebookLM簡報指令.txt",
            mime="text/plain",
            key=f"prompt_tab_detail_{current_week}"
        )
        st.markdown('</div>', unsafe_allow_html=True)
        
        with st.expander("展開直接在網頁複製指令"):
            st.code(prompt_text, language="text")

    # ---- 頁籤三：教學示範 ----
    with tab_demo:
        if current_data['youtube_url']:
            st.markdown('<div class="video-card">', unsafe_allow_html=True)
            st.markdown("### 課程教學示範影片")
            st.write("服事者已上傳之教學錄影。點擊下方按鈕可在「新視窗」開啟觀看，或直接在網頁上播放：")
            
            # 1. 點擊開新視窗按鈕
            st.link_button("在 YouTube 新視窗開啟影片", current_data['youtube_url'])
            
            # 2. 嵌入式網頁播放器
            st.markdown("---")
            st.video(current_data['youtube_url'])
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.info("本課尚未填入教學錄影網址。只要你把 YouTube 連結貼給我，我就會幫你把播放器做出來！")

    st.markdown("---")
    st.caption("召會兒童教材生產線 • 符合主恢復聖徒操練、無宗教肖像原則")
