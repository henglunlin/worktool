import streamlit as st

# 設定主頁的網頁標題與 Layout
st.set_page_config(page_title="Alan 上班偷懶小程式", page_icon="🦼", layout="wide")

# 主頁標題
st.title("🚀 Alan 上班偷懶小程式")
st.markdown("歡迎來到專屬的自動化工具站！請從 **左側邊欄** 選擇你要使用的工具：")

st.markdown("### 🛠️ 工具列表")
st.markdown("* **EDBOM 自動化工具**：將 `EDlist` 與 `BOM` 比對，缺料自動從 `RLClist` 補齊。")
st.markdown("* **BOM 比對工具**：比對新舊 BOM 表差異，並告訴我你會買日月光。")
st.markdown("* **losstool**：補正desense test loss 編輯用。")
st.markdown("* **SNP File Viewer**：比對SNP file 並且能轉換成`Excel file`，並告訴我你會買日月光。")

st.info("👈 點擊左側選單開始偷懶（X）提升效率（O）！祝您鑽大錢！")
