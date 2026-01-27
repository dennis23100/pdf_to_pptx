# 📄 PDF/圖片 轉 PPTX 可編輯工具

將 PDF 或圖片轉換為 PowerPoint，**文字可編輯、可移動**！

## ✨ 功能特點

- 🔍 **OCR 文字偵測** - 自動識別圖片中的文字
- ✏️ **文字可編輯** - 轉換後的文字方塊可以編輯
- 🎨 **智能背景處理** - 自動匹配周圍背景色覆蓋原文字
- 🌐 **多語言支援** - 支援繁體中文、簡體中文、英文、日文等
- 💻 **兩種使用方式** - 網頁版 + Python 本地版

## 🚀 快速開始

### 方式一：網頁版（推薦）

直接訪問：**https://dennis23100.github.io/pdf_to_pptx/**

或者本地開啟 `web/index.html`

### 方式二：Python 本地版

```bash
# 1. 安裝依賴
pip install -r requirements.txt

# 2. 安裝 Tesseract OCR
# Windows: 下載 https://github.com/UB-Mannheim/tesseract/wiki
# Mac: brew install tesseract tesseract-lang
# Linux: sudo apt install tesseract-ocr tesseract-ocr-chi-tra

# 3. 執行
python python/pdf_to_pptx.py 輸入檔案.pdf 輸出檔案.pptx
```

## 📖 使用說明

### 網頁版
1. 開啟網頁
2. 選擇 PDF 或圖片檔案
3. 選擇語言（預設：繁體中文+英文）
4. 點擊「轉換」
5. 下載生成的 PPTX

### Python 版

```bash
# 基本用法
python python/pdf_to_pptx.py input.pdf output.pptx

# 指定語言
python python/pdf_to_pptx.py input.png output.pptx --lang chi_tra+eng

# 生成預覽圖
python python/pdf_to_pptx.py input.pdf output.pptx --preview

# 查看幫助
python python/pdf_to_pptx.py --help
```

### 支援的語言代碼

| 語言 | 代碼 |
|------|------|
| 繁體中文 | `chi_tra` |
| 簡體中文 | `chi_sim` |
| 英文 | `eng` |
| 日文 | `jpn` |
| 韓文 | `kor` |

多語言組合：`chi_tra+eng`（繁中+英文）

## 🛠️ 技術架構

### 網頁版
- **PDF.js** - PDF 渲染
- **Tesseract.js** - 瀏覽器端 OCR
- **PptxGenJS** - PPTX 生成

### Python 版
- **Tesseract OCR** - 文字偵測
- **OpenCV** - 圖像處理
- **python-pptx** - PPTX 生成
- **pdf2image** - PDF 轉圖片

## 📁 專案結構

```
pdf-to-pptx-editable/
├── README.md
├── requirements.txt
├── web/
│   └── index.html      # 網頁版（單檔案）
├── python/
│   └── pdf_to_pptx.py  # Python 版
└── examples/
    └── sample.png      # 範例圖片
```

## ⚠️ 注意事項

1. **OCR 準確度**：辨識準確度取決於圖片品質和字體
2. **複雜背景**：設計類圖片的文字移除效果可能不完美
3. **字體**：轉換後預設使用「Microsoft JhengHei」（微軟正黑體）

## 📝 授權

MIT License - 可自由使用、修改、分發

## 🤝 貢獻

歡迎提交 Issue 或 Pull Request！

---

Made with ❤️ by Claude
