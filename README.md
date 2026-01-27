# 📄 PDF to PPTX Converter

免費開源的 PDF 轉 PowerPoint 工具，保留原始設計，文字方塊可編輯移動。

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![GitHub Pages](https://img.shields.io/badge/GitHub-Pages-green.svg)](https://pages.github.com/)

## ✨ 功能特色

- 🎨 **保留原始設計**：每頁 PDF 完整呈現為投影片背景
- 📝 **文字可編輯**：在 PowerPoint 中可以移動、編輯、刪除文字方塊
- 🔍 **雙模式支援**：
  - **原生文字模式**：提取 PDF 內建文字層（速度快）
  - **OCR 辨識模式**：AI 辨識圖片中的文字（適合掃描版）
- 🌐 **多語言支援**：繁體中文、簡體中文、英文、日文、韓文
- 🔒 **100% 隱私**：完全在瀏覽器中處理，不上傳任何資料
- 📱 **響應式設計**：支援桌面和手機瀏覽器

## 🚀 線上使用

直接訪問：**https://你的用戶名.github.io/pdf-to-pptx/**

## 📁 專案結構

```
pdf-to-pptx/
├── index.html          # 主要網頁介面
├── css/
│   └── style.css       # 樣式表
├── js/
│   ├── app.js          # 主要應用邏輯
│   ├── pdf-parser.js   # PDF 解析模組
│   └── pptx-builder.js # PPTX 建構模組
├── README.md           # 說明文件
└── LICENSE             # MIT 授權
```

## 🛠️ 技術架構

本工具使用以下開源函式庫：

| 函式庫 | 用途 | 版本 |
|--------|------|------|
| [PDF.js](https://mozilla.github.io/pdf.js/) | Mozilla 的 PDF 解析引擎 | 3.11.174 |
| [Tesseract.js](https://tesseract.projectnaptha.com/) | 瀏覽器端 OCR 引擎 | 5.x |
| [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) | PowerPoint 檔案產生器 | 3.12.0 |

### 核心流程

```
PDF 檔案
    ↓
┌───────────────────────────────┐
│  PDF.js 解析                  │
│  - 提取文字內容和位置座標      │
│  - 渲染頁面為高解析度圖片      │
└───────────────────────────────┘
    ↓
┌───────────────────────────────┐
│  文字處理（二選一）            │
│  A. 原生文字：直接使用座標     │
│  B. OCR 模式：Tesseract 辨識  │
└───────────────────────────────┘
    ↓
┌───────────────────────────────┐
│  PptxGenJS 建構 PPTX          │
│  - 背景圖片層                 │
│  - 透明可編輯文字方塊層        │
└───────────────────────────────┘
    ↓
下載 PPTX 檔案
```

## 📖 使用方式

### 基本使用

1. 開啟網站
2. 點擊或拖曳 PDF 檔案到上傳區
3. 選擇轉換模式：
   - **原生文字模式**：適合一般 PDF（有文字層）
   - **OCR 辨識模式**：適合掃描版或設計類 PDF
4. 點擊「開始轉換」
5. 等待完成，檔案會自動下載

### 在 PowerPoint 中編輯

轉換後的 PPTX 包含兩層：

1. **背景層**：原始 PDF 頁面的圖片
2. **文字層**：透明的可編輯文字方塊

你可以：
- 點選文字方塊並移動位置
- 雙擊編輯文字內容
- 刪除不需要的文字方塊
- 添加新的文字方塊

## 🖥️ 本地部署

### 方法一：直接開啟

1. 下載或克隆此專案
2. 用瀏覽器開啟 `index.html`

### 方法二：使用本地伺服器

```bash
# 使用 Python
python -m http.server 8000

# 或使用 Node.js
npx serve .
```

然後訪問 `http://localhost:8000`

## 🌐 部署到 GitHub Pages

1. Fork 或創建此專案的 Repository
2. 進入 Settings → Pages
3. Source 選擇 `main` branch 和 `/ (root)`
4. 點擊 Save
5. 等待幾分鐘後即可訪問

## ⚠️ 注意事項

- **檔案大小**：建議 PDF 不超過 50 頁或 50MB
- **OCR 準確度**：取決於圖片品質，複雜字體可能辨識不完美
- **首次 OCR**：需要下載語言模型（約 10-30MB），請稍等
- **瀏覽器支援**：建議使用 Chrome、Firefox、Edge 最新版

## 🔧 自訂開發

### 修改樣式

編輯 `css/style.css` 來自訂外觀。

### 修改轉換邏輯

- `js/pdf-parser.js`：PDF 解析相關
- `js/pptx-builder.js`：PPTX 建構相關
- `js/app.js`：主要應用邏輯

### 添加新語言

在 `index.html` 的語言選擇器中添加：

```html
<option value="語言代碼">語言名稱</option>
```

語言代碼可參考：[Tesseract 支援的語言](https://tesseract-ocr.github.io/tessdoc/Data-Files-in-different-versions.html)

## 📄 授權

MIT License - 可自由使用、修改、分發

## 🙏 致謝

- [Mozilla PDF.js](https://mozilla.github.io/pdf.js/)
- [Tesseract.js](https://tesseract.projectnaptha.com/)
- [PptxGenJS](https://gitbrent.github.io/PptxGenJS/)

---

Made with ❤️ | 歡迎提交 Issue 和 PR
