import streamlit as st
import io
import time
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# 設置網頁標題與風格
st.set_page_config(page_title="兒童教材自動化生產線", page_icon="📚", layout="centered")

st.title("📚 兒童教材自動化生產線")
st.subheader("用 AI 打造一鍵式教學素材生成平台")
st.write("不需手動複製、不需重複微調，選擇主題即可一鍵匯出繁體教材與簡報。")

st.markdown("---")

# 建立選擇區
col1, col2 = st.columns(2)

with col1:
    week_option = st.selectbox(
        "請選擇課程週次",
        ["第 1 週：認識小動物", "第 2 週：太空探險"]
    )

with col2:
    type_option = st.selectbox(
        "請選擇教材類型",
        ["繪本故事", "科學探索"]
    )

st.markdown("---")

# 模擬不同選擇的教材內容資料庫（台灣繁體中文與符合使用者設定的無特定神祇/十字架意象）
DATA_DB = {
    "第 1 週：認識小動物": {
        "繪本故事": {
            "title": "小貓咪的森林冒險",
            "content": "在一個陽光普照的早晨，小貓咪奇奇決定去森林裡探險。一路上，奇奇遇到了嗡嗡叫的小蜜蜂、在樹梢上唱歌的小鳥，還有正在搬運橡實的小松鼠。大家雖然長得不一樣，但都非常友善，奇奇學到了「分享」與「互助」的快樂，最後帶著滿滿的收穫，在溫暖的光束照引下平安回到了家。",
            "scenes": [
                "分鏡 1：小貓咪奇奇在陽光下伸懶腰，準備出發探險。",
                "分鏡 2：奇奇在花叢間遇到小蜜蜂，學習蜜蜂採蜜的辛勞。",
                "分鏡 3：奇奇在樹下看到小松鼠們互相幫忙搬橡實，體會合作的重要。",
                "分鏡 4：傍晚時分，天空出現一道溫暖的光束，指引奇奇平安回到溫馨的家。"
            ],
            "notebooklm": "請根據以下結構，為『小貓咪的森林冒險』兒童故事生成一份精美的導讀簡報：\n1. 封面：故事主題與核心精神（分享與互助）\n2. 角色介紹：小貓咪奇奇與森林裡的動物朋友們\n3. 劇情發展：從出發、遇到困難到互相幫忙\n4. 啟發與總結。請注意：簡報視覺設計請採用溫暖活潑色調，改以光束替代具體神聖形象，切勿出現任何十字架、雙手合十或宗教符號。"
        },
        "科學探索": {
            "title": "動物們的超級感官",
            "content": "你知道嗎？大自然裡的動物都擁有像超能力一樣的感官！狗狗的嗅覺比人類靈敏一萬倍以上；貓咪在黑夜裡也能看得一清二楚；蝙蝠則是用我們聽不到的超音波來辨識方向。這些神奇的感官幫助牠們在自然環境中生活與繁衍，是大自然最精妙的設計。",
            "scenes": [
                "分鏡 1：主題封面 - 動物超級感官大解密。",
                "分鏡 2：嗅覺篇 - 介紹狗狗強大的嗅覺與嗅覺細胞。",
                "分鏡 3：視覺篇 - 介紹貓咪視網膜的特殊結構與夜視能力。",
                "分鏡 4：聽覺與回聲篇 - 介紹蝙蝠如何利用超音波在黑暗中自由飛翔。"
            ],
            "notebooklm": "請將『動物們的超級感官』科學教材轉化為生動的教學投影片大綱。包含：\n1. 引言：什麼是超級感官？\n2. 單元一：嗅覺冠軍（犬類）\n3. 單元二：夜視高手（貓科）\n4. 單元三：回聲定位（蝙蝠）。請提供互動式的課堂提問設計，視覺風格採科技與自然風格，不要包含任何宗教或特定圖騰。"
        }
    },
    "第 2 週：太空探險": {
        "繪本故事": {
            "title": "小星星的尋家之旅",
            "content": "夜空裡有一顆迷路的小星星叫做亮亮。亮亮在浩瀚的宇宙中穿梭，拜訪了巨大的木星、擁有漂亮光環的土星，還有紅通通的火星。在旅行中，亮亮發現每顆行星都有自己獨特的美麗。最後，在宇宙星雲散發出的溫柔光束擁抱下，亮亮找到了最適合自己的軌道，安心地閃爍著光芒。",
            "scenes": [
                "分鏡 1：小星星亮亮迷路了，在黑暗的宇宙中顯得有些孤單。",
                "分鏡 2：亮亮拜訪木星與土星，欣賞牠們巨大且美麗的外觀。",
                "分鏡 3：亮亮與火星聊天，了解不同星球的自然環境。",
                "分鏡 4：星雲化作溫柔的光束包裹著亮亮，亮亮終於找到自己的定位，快樂地發光。"
            ],
            "notebooklm": "請為兒童繪本『小星星的尋家之旅』設計課程簡報腳本：\n1. 引人入勝的開頭：迷路的小星星\n2. 行星巡禮：木星、土星、火星的特色\n3. 寓意傳達：每個人都是獨一無二的存在\n4. 結尾。視覺提示：主色調為深藍色宇宙，神聖或溫馨場景請用純淨光束表達，避免任何宗教神明、翅膀或雙手合十的圖像。"
        },
        "科學探索": {
            "title": "神秘的太陽系八大行星",
            "content": "太陽系是一個以太陽為中心的巨大大家庭，圍繞著它運行的有八大行星。離太陽最近的水星非常炙熱，而最遠的海王星則是一顆冰冷藍色星球。我們居住的地球，因為有適宜的水分、空氣與溫度，成為了目前已知唯一擁有豐富生命的藍色綠洲。讓我們一起探索這趟宇宙奧秘之旅吧！",
            "scenes": [
                "分鏡 1：太陽系全景導覽 - 八大行星的排列位置與比例。",
                "分鏡 2：內太陽系 - 水星、金星、地球與火星的岩石結構。",
                "分鏡 3：外太陽系 - 木星、土星、天王星與海王星的氣體與冰雪世界。",
                "分鏡 4：生命的奇蹟 - 為什麼地球是目前最特別的生命綠洲？"
            ],
            "notebooklm": "請幫我將『神秘的太陽系八大行星』科學文本優化為適合 NotebookLM 的語音導讀與簡報指令。要求羅列出各行星的核心科學數據（體積、距離、特點），並設計一個5頁的精簡簡報架構，風格偏向現代天文科普，視覺乾淨，不包含任何十字架、天使或特定神明繪圖。"
        }
    }
}

# 輔助函式：生成 Word 檔
def generate_docx(title, content, scenes):
    doc = Document()
    doc.add_heading(title, level=1)
    
    doc.add_heading("【教材核心內容】", level=2)
    doc.add_paragraph(content)
    
    doc.add_heading("【課堂教學分鏡/分頁建議】", level=2)
    for scene in scenes:
        doc.add_paragraph(scene, style='List Bullet')
        
    doc.add_paragraph("\n* 本教材已自動優化為台灣繁體中文語系，並自動過濾敏感意象。")
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# 輔助函式：生成 PPT 檔
def generate_pptx(title, scenes):
    prs = Presentation()
    # 使用空白版面
    blank_slide_layout = prs.slide_layouts[6]
    
    # 第一頁：封面
    slide = prs.slides.add_slide(blank_slide_layout)
    txBox = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(2))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = RGBColor(15, 23, 42)
    
    p2 = tf.add_paragraph()
    p2.text = "每週每課精選簡報"
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(100, 116, 139)
    
    # 後續分鏡頁
    for i, scene in enumerate(scenes, 1):
        slide = prs.slides.add_slide(blank_slide_layout)
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
        tf = txBox.text_frame
        
        p_title = tf.paragraphs[0]
        p_title.text = f"第 {i} 頁：課程焦點"
        p_title.font.size = Pt(28)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(56, 189, 248)
        
        p_content = tf.add_paragraph()
        p_content.text = scene
        p_content.font.size = Pt(20)
        p_content.font.color.rgb = RGBColor(51, 65, 85)
        
    file_stream = io.BytesIO()
    prs.save(file_stream)
    file_stream.seek(0)
    return file_stream

# 一鍵生成邏輯
if st.button("🚀 開始一鍵生成", type="primary"):
    with st.spinner("AI 正在解析主題、自動翻譯台灣繁體中文並建構分鏡中..."):
        time.sleep(1.5) # 模擬 AI 運算時間
        
        # 獲取對應資料
        data = DATA_DB[week_option][type_option]
        
        st.success("✨ 一鍵自動化生成完畢！")
        
        # 顯示成果
        st.markdown(f"### 📋 預覽：{data['title']} ({type_option})")
        st.write(data['content'])
        
        st.markdown("#### 🎬 智能分鏡腳本")
        for s in data['scenes']:
            st.markdown(f"- {s}")
            
        st.markdown("#### 🤖 NotebookLM 簡報專用指令 (Prompt)")
        st.info(data['notebooklm'])
        
        # 將檔案存在 session_state 中供下載按鈕使用
        st.session_state['docx_file'] = generate_docx(data['title'], data['content'], data['scenes'])
        st.session_state['pptx_file'] = generate_pptx(data['title'], data['scenes'])
        st.session_state['file_title'] = data['title']
        st.session_state['generated'] = True

# 下載區域
if st.session_state.get('generated', False):
    st.markdown("---")
    st.markdown("### 📥 下載專區")
    
    col_dl1, col_dl2 = st.columns(2)
    
    with col_dl1:
        st.download_button(
            label="📄 下載 Word 教材 (.docx)",
            data=st.session_state['docx_file'],
            file_name=f"{st.session_state['file_title']}_教材.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    with col_dl2:
        st.download_button(
            label="📊 下載 PPT 簡報 (.pptx)",
            data=st.session_state['pptx_file'],
            file_name=f"{st.session_state['file_title']}_簡報.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
