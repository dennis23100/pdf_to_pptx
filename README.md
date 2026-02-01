# 📄 PDF 轉 PPTX 專業版

[![GitHub Pages](https://img.shields.io/badge/Demo-GitHub%20Pages-blue)](https://你的用戶名.github.io/pdf-to-pptx-pro/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

將 PDF 或圖片轉換為 PowerPoint，**文字可編輯、可移動**！


## ✨ 特色功能

| 功能 | 說明 |
|------|------|
| 🤖 **AI 文字移除** | 使用 LaMa ONNX 模型在瀏覽器端移除圖片中的文字 |
| 🔍 **OCR 文字識別** | 使用 Tesseract.js 識別文字位置和內容 |
| ✏️ **文字可編輯** | 轉換後的 PPTX 文字方塊可以自由編輯 |
| 🆓 **完全免費** | 不需要任何 API Key，無使用限制 |
| 🔒 **隱私保護** | 所有處理在瀏覽器本地完成，檔案不上傳 |
| 🌐 **多語言 OCR** | 支援繁體中文、簡體中文、英文、日文、韓文 |

## 🚀 線上使用

部署後訪問：`https://你的用戶名.github.io/pdf-to-pptx-pro/`

## 🎯 三種處理模式

### 1. AI 文字移除（推薦）
- 使用 LaMa AI 模型移除圖片中的文字
- 生成乾淨的背景 + 可編輯文字層
- 效果最好，處理較慢
- 首次使用需下載 ~50MB AI 模型

### 2. 背景色覆蓋（快速）
- 用背景色覆蓋原文字區域
- 處理速度快，無需下載模型
- 適合簡單背景的文件

### 3. 純圖片模式
- 直接將頁面轉為圖片
- 最快速，但文字不可編輯

## 🛠️ 技術架構

```
┌─────────────────────────────────────────────────────────────┐
│                      技術棧                                  │
├─────────────────────────────────────────────────────────────┤
│  PDF.js          │ PDF 解析和渲染                           │
│  Tesseract.js    │ 瀏覽器端 OCR 文字識別                    │
│  LaMa ONNX       │ AI 文字移除（瀏覽器端運行）              │
│  ONNX Runtime    │ 在瀏覽器中運行 AI 模型                   │
│  PptxGenJS       │ 生成 PowerPoint 檔案                     │
└─────────────────────────────────────────────────────────────┘
```

## 📁 專案結構

```
pdf-to-pptx-pro/
├── index.html     # 主頁面
├── app.js         # 核心應用程式
├── README.md      # 說明文件
└── LICENSE        # MIT 授權
```

## 🚀 部署到 GitHub Pages

### 步驟 1：創建 Repository

1. 前往 https://github.com/new
2. Repository name：`pdf-to-pptx-pro`
3. 選擇 **Public**
4. 點擊 **Create repository**

### 步驟 2：上傳檔案

1. 在 repo 頁面點擊 **uploading an existing file**
2. 上傳所有檔案：
   - `index.html`
   - `app.js`
   - `README.md`
   - `LICENSE`
3. 點擊 **Commit changes**

### 步驟 3：啟用 GitHub Pages

1. 進入 **Settings** → **Pages**
2. Source 選擇 `Deploy from a branch`
3. Branch 選擇 `main`，資料夾選 `/ (root)`
4. 點擊 **Save**

### 步驟 4：訪問

等待 1-2 分鐘後：

```
https://你的用戶名.github.io/pdf-to-pptx-pro/
```

## 📖 使用說明

1. **上傳檔案** - 點擊或拖拽 PDF/圖片
2. **選擇設定** - 語言、處理模式、品質
3. **選擇頁面** - 勾選要轉換的頁面
4. **開始轉換** - 等待 AI 處理
5. **下載結果** - PPTX 自動下載

## ⚠️ 注意事項

### AI 模式
- 首次使用需下載約 50MB 的 AI 模型
- 模型會緩存在瀏覽器中，之後無需重新下載
- 處理時間取決於電腦性能（每頁約 5-15 秒）

### 瀏覽器支援
- 推薦：Chrome 90+、Edge 90+
- 支援：Firefox 90+、Safari 15+
- 需要支援 WebAssembly

### 限制
- 單檔最大 100MB
- 複雜背景的文字移除效果可能不完美
- 大量頁面處理時可能較慢

## 🔧 本地開發

```bash
# 克隆專案
git clone https://github.com/你的用戶名/pdf-to-pptx-pro.git
cd pdf-to-pptx-pro

# 使用任何 HTTP 伺服器
python -m http.server 8080
# 或
npx serve
```

## 📊 與其他工具對比

| 功能 | 本專案 | NBLM2PPTX | DeckEdit |
|------|--------|-----------|----------|
| 免費使用 | ✅ 完全免費 | ⚠️ 需 API Key | ✅ 免費 |
| 無限使用 | ✅ 無限制 | ❌ 有配額限制 | ✅ 無限制 |
| AI 文字移除 | ✅ LaMa ONNX | ✅ Gemini AI | ❓ 不確定 |
| 本地處理 | ✅ 完全本地 | ⚠️ 需調用 API | ✅ 本地 |
| 開源 | ✅ MIT | ✅ MIT | ❌ 閉源 |

## 🙏 致謝

- [PDF.js](https://mozilla.github.io/pdf.js/) - Mozilla
- [Tesseract.js](https://tesseract.projectnaptha.com/) - Project Naptha
- [LaMa](https://github.com/advimman/lama) - SAIC-Moscow
- [LaMa ONNX](https://huggingface.co/Carve/LaMa-ONNX) - Carve Photos
- [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) - Brent Ely
- [ONNX Runtime Web](https://onnxruntime.ai/) - Microsoft

## 📝 License

MIT License - 可自由使用、修改、分發

---

Made with ❤️
