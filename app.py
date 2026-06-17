import streamlit as st
import io
from docx import Document

# 設置網頁風格與標題
st.set_page_config(page_title="召會兒童班教材自動化生產線 v3.8", page_icon="💡", layout="centered")

# 注入更活潑、針對頁籤與既有圖片連結優化的 CSS 視覺
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stButton>button { width: 100%; border-radius: 12px; height: 3.5em; font-weight: bold; font-size: 16px; transition: 0.3s; }
    .stButton>button:hover { transform: scale(1.02); }
    .download-card { padding: 24px; border-radius: 16px; background: white; border: 1px solid #e2e8f0; margin-bottom: 20px; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05); }
    .video-card { padding: 24px; border-radius: 16px; background: #f0fdf4; border: 1px solid #bbf7d0; margin-bottom: 20px; }
    .slides-card { padding: 24px; border-radius: 16px; background: #eff6ff; border: 1px solid #bfdbfe; margin-bottom: 20px; }
    h1 { color: #0f172a; font-weight: 800; }
    h3 { color: #1e293b; font-weight: 700; margin-top: 10px; }
    /* 頁籤字體放大且顯眼 */
    .stTabs [data-baseweb="tab"] { font-size: 18px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("📚 召會兒童教材自動化生產線 3.8")
st.subheader("一鍵匯出繁體教材 Word 與 NotebookLM 簡報指令碼")
st.write("本系統已完全優化為台灣繁體中文，融入主恢復編輯習慣（轉向靈、呼求主名、生命供應），並全面過濾不適宜宗教意象，改以溫暖的光束指引。")

st.markdown("---")

# --- 1 至 40 課大型真實資料庫建構 ---
CURRICULUM_DB = {
    1: {
        "title": "第 1 週：兒女是神賜給父母的產業",
        "hymn": "兒童詩歌第 194 頁《當我還是個小小嬰孩》",
        "verse": "「兒女是耶和華所賜的產業；所懷的胎，是他所給的賞賜。」— 詩篇一百二十七篇 3 節",
        "discussion": "1. 請小朋友簡單介紹並描述自己的父母。\\n2. 討論我們是怎麼來的，我們的父母又是怎麼來的。\\n3. 討論人為什麼要結婚？人結婚後為什麼要有孩子？",
        "story": "神照著自己的形像和樣式用地上的塵土造了亞當，盼望人在地上彰顯祂並代表祂。神看亞當獨居不好，便在他熟睡時用肋骨造了夏娃作他完美的配偶，使他們在愛與一裡合為一體。神賜福給這第一對夫妻，要他們生養眾多。孩子是神賜給父母的產業與祝福。舊約的哈拿因沒有孩子而在殿裡向耶和華傾心吐意、迫切流淚禱告，神便顧念她，賜下日後偉大的神人撒母耳，除去她的羞恥，使哈拿的心因耶和華歡欣稱頌。",
        "exercise": "本週操練：在禱告中感謝神把我們擺在一個敬虔、和樂的家庭中，使我們能享受神給我們的祝福。",
        "pure_text": "第一課：兒女是神賜給父母的產業\\n\\n【本週背經】\\n「兒女是耶和華所賜的產業；所懷的胎，是他所給的賞賜。」— 詩篇一百二十七篇 3 節\\n「神就賜福給他們，又對他們說，要生養眾多，遍滿地面，治理這地。」— 創世記一章 28 節\\n\\n【教材內容】\\n神照著自己的形像，按著自己的樣式創造了人，盼望人在地上彰顯祂並代表祂。在神最初的創造中，神與人的關係是非常親密且彼此瞭解的。隨後，神使亞當沈睡，取其肋骨造了夏娃作完美的配偶。亞當一見便歡喜說：『這是我骨中的骨，肉中的肉。』兩個人在愛與一裡有著甜美的親密關係。神也賜福這第一對夫妻，要他們生養眾多、遍滿地面。因此，孩子是神給父母的產業與賞賜。在以色列人中，沒有孩子曾是一件羞恥的事，但哈拿在殿裡向耶和華傾心吐意、迫切禱告，神便顧念她並應允她，賜下日後偉大的神人撒母耳。這說明孩子是從神而來的祝福，生命延續的喜樂能使家庭的愛更加甜蜜與擴大。",
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
        "story": "神託付兒女對父母盡孝敬與順從的責任。孝敬的表現就是愛和尊敬父母。
