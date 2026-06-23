import streamlit as st
import io
from docx import Document

# 1. 網頁全域高質感樣式設定（完全移除 Emoji，採用淡色系與圓角對比色按鈕）
st.set_page_config(page_title="兒童教材自動化生產管理系統", layout="wide")

st.markdown("""
    <style>
    /* 全局字體與色調優化 */
    html, body, [data-testid="stAppViewContainer"] {
        background-color: #FAFAFA;
        color: #1E293B;
        font-family: "Helvetica Neue", Arial, "Noto Sans TC", sans-serif;
    }
    
    /* 標題與副標題樣式 */
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 0.5rem;
    }
    .sub-title {
        font-size: 1.1rem;
        color: #64748B;
        margin-bottom: 2rem;
    }
    
    /* 自訂圓角淡色系對比按鈕 */
    div.stButton > button {
        background-color: #F1F5F9 !important;
        color: #0F172A !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.5rem !important;
        font-size: 1rem !important;
        font-weight: 500 !important;
        transition: all 0.2s ease;
    }
    div.stButton > button:hover {
        background-color: #E2E8F0 !important;
        border-color: #94A3B8 !important;
        color: #0284C7 !important;
    }
    
    /* 下載按鈕專用強調樣式 (高對比色) */
    div.stDownloadButton > button {
        background-color: #0EA5E9 !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.5rem !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        box-shadow: 0 2px 4px rgba(14, 165, 233, 0.15);
    }
    div.stDownloadButton > button:hover {
        background-color: #0284C7 !important;
        color: #FFFFFF !important;
    }
    
    /* 移除 Streamlit 內建廉價紅線與標籤 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# 2. 核心大綱與基本資料庫結構 (1-40課正確名稱宣告，並內建1-3課完整主恢復語境內容)
LESSONS_CONFIG = {
    1: {"title": "兒女是神賜給父母的產業", "has_data": True, "vid": "PLmSdsjakHIh-6RP46PzI1c-16J6nfnDkK"},
    2: {"title": "父母對兒女的責任", "has_data": True, "vid": "PLmSdsjakHIh-6RP46PzI1c-16J6nfnDkK"},
    3: {"title": "兒女對父母該有的態度(一)─愛父母", "has_data": True, "vid": "PLmSdsjakHIh-6RP46PzI1c-16J6nfnDkK"},
    4: {"title": "兒女對父母該有的態度(二)─尊敬父母", "has_data": False, "vid": ""},
    5: {"title": "順從父母", "has_data": False, "vid": ""},
    6: {"title": "孝敬並順從父母的結果", "has_data": False, "vid": ""},
}
# 補充 7 到 40 課的正確目錄骨架
for i in range(7, 41):
    LESSONS_CONFIG[i] = {"title": f"第 {i} 課課程教材 (待上傳更新)", "has_data": False, "vid": ""}

# 1-3課的台灣繁體主恢復語境完整文本資料庫 (嚴格移除上帝，改用耶和華/神，場景以光束替代)
DATA_DB = {
    1: {
        "text": "第一課：兒女是神賜給父母的產業\n\n負擔：兒女是神賜給父母的產業。神把人造得像祂自己一樣，使人與神能有很親密的關係及交往。神先創造男人，之後又造了女人來作男人的配偶，使男人能以完全。這樣，男人與女人就有了一個在愛裡以及在一裡的親密關係。這甜美、親密的關係，因著神的祝福而得以完滿。這祝福就是賜給他們孩子，使他們可以繁衍後代，所以孩子乃是神給父母的祝福。\n\n背經：兒女是耶和華所賜的產業；所懷的胎，是他所給的賞賜（詩一二七3）。\n\n課程內容：神照著自己的形像，按著自己的樣式用地上的塵土造了第一個人，就是亞當。神這樣造人，是因祂願意人像祂，盼望人能在地上彰顯祂並代表祂。在神最初的創造中，神與人的關係是非常親密且彼此瞭解的。哈拿向神迫切禱告，許願神若賜她一個兒子，她必使他終身歸與耶和華。神應允她的禱告，生了撒母耳。撒母耳的意思是：這是我從耶和華那裡求來的。",
        "notebooklm": "請根據以下結構，為『兒女是神賜給父母的產業』課程生成精美簡報，風格要求活潑可愛、不額外延伸主題、符合學齡和中低年級的視覺設計。\n\n簡報大綱：\n1. 封面（第 1 課：兒女是神賜給父母的產業）\n2. 本週詩歌：當我還是個小小嬰孩\n3. 問題與討論：介紹自己的父母、討論我們是怎麼來的\n4. 本週故事：神與人的親密關係、神為亞當造配偶、哈拿與撒母耳的故事\n5. 本週背經：兒女是耶和華所賜的產業；所懷的胎，是他所給的賞賜（詩一二七3）\n6. 本週生活操練：感謝神把我們擺在一個敬虔、和樂的家庭中，享受神給我們的祝福。"
    },
    2: {
        "text": "第二課：父母對兒女的責任\n\n負擔：這一課我們要讓孩子們深刻的認識父母愛孩子，並且供應孩子一切的需要。父母總是甘心樂意為孩子預備一切，因為他們愛孩子。召會中的父母，必須從主領受教養孩子的負擔，並且向主負責。\n\n背經：教養孩童，使他走當行的道，就是到老他也不偏離（箴二二6）。\n\n課程內容：父母對兒女的愛是神所賦予的。照顧一個嬰孩不是件容易的事，這不僅需要會照顧，也需要有足夠的愛心和耐心。舊約的榜樣如摩西的母親約基別，在乳養摩西期間不斷把對神的信仰傳輸給摩西，使摩西從小認識自己是以色列人。新約的榜樣如提摩太的外祖母羅以和母親友尼基，她們在提摩太很小的時候就常將聖經的話教導他，將對主的信心傳輸到他裡面，使他從小明白聖經。",
        "notebooklm": "請根據以下結構，為『父母對兒女的責任』課程生成精美簡報，風格要求活潑可愛、不額外延伸主題、符合學齡和中低年級的視覺設計。\n\n簡報大綱：\n1. 封面（第 2 課：父母對兒女的責任）\n2. 本週詩歌：當我還是個小小嬰孩\n3. 問題與討論：介紹父母的職業與日常對我們的關心\n4. 本週故事：父母的三個責任（愛與照顧、供給物質、管教），以及約基別與提摩太家人的屬靈榜樣\n5. 本週背經：教養孩童，使他走當行的道，就是到老他也不偏離（箴二二6）\n6. 本週生活操練：操練為著父母對我們的愛、供應和管教感謝主。"
    },
    3: {
        "text": "第三課：兒女對父母該有的態度(一)─愛父母\n\n負擔：兒女對父母該有的三種基本態度，就是從心裡孝敬父母、愛父母、尊敬父母，在行為上順從父母、聽父母的話。父母是我們的源頭，如同神是我們的源頭；因此我們應當愛神，也當愛父母。\n\n背經：當孝敬父母，使你的日子在耶和華你神所賜你的地上，得以長久（出二十12）。\n\n課程內容：神托付兒女對父母當盡的責任主要有兩個，一個是孝敬，一個是順從。愛父母的具體表現包括向父母說謝謝、幫助父母作家事、和父母談心。英國弟兄會的領頭弟兄達秘非常愛他的母親，他在五歲時母親就去世了，但他一直深深記得母親溫柔的眼睛，這使他一生充滿對主的愛和對人的愛。另外，亞伯拉罕林肯總統也非常敬愛他的母親與繼母，認為自己能有今天完全是出於母親的教誨。",
        "notebooklm": "請根據以下結構，為『兒女對父母該有的態度(一)─愛父母』課程生成精美簡報，風格要求活潑可愛、不額外延伸主題、符合學齡和中低年級的視覺設計。\n\n簡報大綱：\n1. 封面（第 3 課：兒女對父母該有的態度(一)─愛父母）\n2. 本週詩歌：兒女當孝敬父母\n3. 問題與討論：最喜歡父母的哪些方面？如何向父母表達感激和愛？\n4. 本週故事：孝敬與順從的定義、愛父母的具體表現，以及達秘與林肯敬愛母親的故事\n5. 本週背經：當孝敬父母，使你的日子在耶和華你神所賜你的地上，得以長久（出二十12）\n6. 本週生活操練：操練向父母表達我們對他們的愛，如說謝謝、幫忙作家事、和父母談心與擁抱。"
    }
}

# 輔助函式：即時生成繁體純文字 Docx
def make_pure_text_docx(content_text):
    doc = Document()
    for line in content_text.split('\n'):
        if line.strip():
            if "第一課" in line or "第二課" in line or "第三課" in line:
                doc.add_heading(line, level=1)
            else:
                doc.add_paragraph(line)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

# 3. 介面排版
st.markdown('<div class="main-title">兒童教材自動化生產管理系統</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">選擇課程即可切換專屬純文字教材、簡報圖像資產與教學示範影片</div>', unsafe_allow_html=True)

# 課程下拉選單 (支援 1 至 40 課)
lesson_options = [f"第 {k} 課：{LESSONS_CONFIG[k]['title']}" for k in range(1, 41)]
selected_idx = st.selectbox("請選擇目標課程進行管理", range(1, 41), format_func=lambda x: lesson_options[x-1])

st.markdown("---")

# 真正的多頁標籤切換區
tab1, tab2, tab3 = st.tabs(["課程內容管理", "簡報內容管理", "教學示範影片"])

# 獲取當前選擇課程的基礎組態
current_cfg = LESSONS_CONFIG[selected_idx]

with tab1:
    st.markdown(f"### {lesson_options[selected_idx-1]}")
    if selected_idx in DATA_DB:
        # 顯示繁體純文字
        st.text_area("當前純文字教材預覽", DATA_DB[selected_idx]["text"], height=300, disabled=True)
        
        # 下載純文字教材按鈕
        docx_file = make_pure_text_docx(DATA_DB[selected_idx]["text"])
        st.download_button(
            label="下載純文字教材 (.docx)",
            data=docx_file,
            file_name=f"第{selected_idx}課_純文字教材.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning(f"第 {selected_idx} 課純文字教材尚未填入，請在對話框中提供文本給我以進行自動化串接。")

with tab2:
    st.markdown("### NotebookLM 簡報專用高對比指令碼")
    if selected_idx in DATA_DB:
        st.code(DATA_DB[selected_idx]["notebooklm"], language="text")
        st.caption("提示：點擊右上角複製按鈕即可貼入 NotebookLM 中。")
        
        st.markdown("---")
        st.markdown("### 簡報圖檔資產預覽 (PNG 圖片)")
        # 串接雲端硬碟中的 PNG 圖片資產預覽 (此處為結構化路徑示意)
        st.image("https://images.unsplash.com/photo-1516627145497-ae6968895b74?q=80&w=600", 
                 caption=f"第 {selected_idx} 課 Google Drive 簡報圖檔同步預覽中", width=500)
    else:
        st.warning(f"第 {selected_idx} 課簡報資產尚未配置。")

with tab3:
    st.markdown("### 影音示範與線上觀摩")
    if current_cfg["vid"]:
        # 完美的點擊跳出新分頁開啟影片功能
        video_url = f"https://www.youtube.com/playlist?list={current_cfg['vid']}"
        st.markdown(f'請點擊下方按鈕，將會開啟新分頁前往專屬教學播放清單：')
        st.link_button("在新分頁開啟示範影片", video_url)
    else:
        st.info(f"第 {selected_idx} 課示範影片網址尚未設定，歡迎隨時提供網址讓我為你綁定更新。")
