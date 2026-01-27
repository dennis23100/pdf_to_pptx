/**
 * PDF to PPTX Converter - Main Application
 * 主要應用程式邏輯
 */

const App = {
  // 狀態
  state: {
    file: null,
    mode: 'native',  // 'native' 或 'ocr'
    language: 'chi_tra+eng',
    isConverting: false,
    tesseractWorker: null
  },

  // DOM 元素快取
  elements: {},

  /**
   * 初始化應用程式
   */
  init() {
    // 初始化 PDF.js Worker
    pdfjsLib.GlobalWorkerOptions.workerSrc = 
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

    // 快取 DOM 元素
    this.cacheElements();
    
    // 綁定事件
    this.bindEvents();

    console.log('PDF to PPTX Converter initialized');
  },

  /**
   * 快取 DOM 元素
   */
  cacheElements() {
    this.elements = {
      fileInput: document.getElementById('fileInput'),
      uploadZone: document.getElementById('uploadZone'),
      uploadPrompt: document.getElementById('uploadPrompt'),
      fileInfo: document.getElementById('fileInfo'),
      fileName: document.getElementById('fileName'),
      fileMeta: document.getElementById('fileMeta'),
      modeOptions: document.querySelectorAll('.option-card[data-mode]'),
      langSelect: document.getElementById('langSelect'),
      convertBtn: document.getElementById('convertBtn'),
      resetBtn: document.getElementById('resetBtn'),
      progressSection: document.getElementById('progressSection'),
      progressFill: document.getElementById('progressFill'),
      progressText: document.getElementById('progressText'),
      progressDetail: document.getElementById('progressDetail'),
      successBox: document.getElementById('successBox'),
      errorBox: document.getElementById('errorBox'),
      previewSection: document.getElementById('previewSection'),
      previewContainer: document.getElementById('previewContainer')
    };
  },

  /**
   * 綁定事件處理器
   */
  bindEvents() {
    const { fileInput, uploadZone, modeOptions, convertBtn, resetBtn } = this.elements;

    // 檔案選擇
    fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

    // 拖曳上傳
    uploadZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      uploadZone.classList.add('dragover');
    });

    uploadZone.addEventListener('dragleave', () => {
      uploadZone.classList.remove('dragover');
    });

    uploadZone.addEventListener('drop', (e) => {
      e.preventDefault();
      uploadZone.classList.remove('dragover');
      const file = e.dataTransfer.files[0];
      if (file && file.type === 'application/pdf') {
        this.setFile(file);
      } else {
        this.showError('請選擇 PDF 檔案');
      }
    });

    // 模式選擇
    modeOptions.forEach(option => {
      option.addEventListener('click', () => {
        modeOptions.forEach(o => o.classList.remove('active'));
        option.classList.add('active');
        this.state.mode = option.dataset.mode;
      });
    });

    // 語言選擇
    if (this.elements.langSelect) {
      this.elements.langSelect.addEventListener('change', (e) => {
        this.state.language = e.target.value;
      });
    }

    // 轉換按鈕
    convertBtn.addEventListener('click', () => this.startConversion());

    // 重置按鈕
    resetBtn.addEventListener('click', () => this.reset());
  },

  /**
   * 處理檔案選擇
   */
  handleFileSelect(e) {
    const file = e.target.files[0];
    if (file && file.type === 'application/pdf') {
      this.setFile(file);
    }
  },

  /**
   * 設定檔案
   */
  setFile(file) {
    this.state.file = file;
    this.showFileInfo(file);
    this.loadPreview(file);
    this.elements.convertBtn.disabled = false;
    this.hideError();
  },

  /**
   * 顯示檔案資訊
   */
  showFileInfo(file) {
    const { uploadZone, uploadPrompt, fileInfo, fileName, fileMeta } = this.elements;
    
    uploadPrompt.style.display = 'none';
    fileInfo.style.display = 'block';
    fileInfo.classList.add('show');
    uploadZone.classList.add('has-file');
    
    fileName.textContent = file.name;
    fileMeta.textContent = this.formatBytes(file.size);
  },

  /**
   * 載入預覽
   */
  async loadPreview(file) {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      
      const { previewSection, previewContainer } = this.elements;
      previewContainer.innerHTML = '';
      
      const maxPreview = Math.min(pdf.numPages, 4);

      for (let i = 1; i <= maxPreview; i++) {
        const page = await pdf.getPage(i);
        const scale = 0.4;
        const viewport = page.getViewport({ scale });
        
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');
        
        await page.render({ canvasContext: ctx, viewport }).promise;
        
        const div = document.createElement('div');
        div.className = 'preview-page';
        
        const img = document.createElement('img');
        img.src = canvas.toDataURL('image/jpeg', 0.6);
        img.alt = `第 ${i} 頁`;
        
        const num = document.createElement('div');
        num.className = 'preview-page-num';
        num.textContent = `第 ${i} 頁`;
        
        if (i === maxPreview && pdf.numPages > maxPreview) {
          num.textContent += ` (共 ${pdf.numPages} 頁)`;
        }
        
        div.appendChild(img);
        div.appendChild(num);
        previewContainer.appendChild(div);
      }
      
      previewSection.classList.add('show');
    } catch (err) {
      console.error('Preview error:', err);
    }
  },

  /**
   * 開始轉換
   */
  async startConversion() {
    if (!this.state.file || this.state.isConverting) return;

    this.state.isConverting = true;
    this.updateUI('converting');
    this.hideError();

    try {
      // 1. 讀取 PDF
      this.updateProgress(5, '讀取 PDF...', '');
      const arrayBuffer = await this.state.file.arrayBuffer();

      // 2. 解析 PDF
      const parsedPDF = await PDFParser.parse(arrayBuffer, (progress) => {
        const pct = 5 + (progress.current / progress.total) * 30;
        this.updateProgress(pct, progress.message, '');
      });

      // 3. 處理文字（原生或 OCR）
      const processedPages = await this.processText(parsedPDF);

      // 4. 建構 PPTX
      const pptx = await PPTXBuilder.build(processedPages, {
        addTextBoxes: true,
        textTransparency: 100,  // 透明文字，可編輯
        fontFace: 'Microsoft JhengHei',
        showPageNumber: false
      }, (progress) => {
        const pct = 70 + (progress.current / progress.total) * 25;
        this.updateProgress(pct, progress.message, '');
      });

      // 5. 下載檔案
      this.updateProgress(98, '產生檔案...', '');
      const outputName = this.state.file.name.replace('.pdf', '_editable.pptx');
      await PPTXBuilder.download(pptx, outputName);

      // 完成
      this.updateProgress(100, '完成！', '');
      this.updateUI('success');

    } catch (err) {
      console.error('Conversion error:', err);
      this.showError('轉換失敗: ' + (err.message || '未知錯誤'));
      this.updateUI('error');
    } finally {
      this.state.isConverting = false;
      await this.cleanupOCR();
    }
  },

  /**
   * 處理文字（原生提取或 OCR）
   */
  async processText(parsedPDF) {
    const pages = parsedPDF.pages;
    const processedPages = [];

    for (let i = 0; i < pages.length; i++) {
      const page = pages[i];
      let textLines = [];

      // 檢查是否有原生文字
      if (page.hasNativeText && this.state.mode === 'native') {
        // 使用原生文字
        textLines = PDFParser.groupTextByLines(page.textItems, page.height * 0.02);
        
        this.updateProgress(
          35 + (i / pages.length) * 35,
          `處理第 ${i + 1}/${pages.length} 頁`,
          '提取原生文字'
        );
      } else if (this.state.mode === 'ocr') {
        // 使用 OCR
        textLines = await this.performOCR(page, i, pages.length);
      }

      processedPages.push({
        ...page,
        textLines: textLines
      });
    }

    return processedPages;
  },

  /**
   * 執行 OCR
   */
  async performOCR(page, index, total) {
    // 初始化 Tesseract Worker（如果尚未初始化）
    if (!this.state.tesseractWorker) {
      this.updateProgress(35, '初始化 OCR 引擎...', '首次使用需下載語言模型');
      
      this.state.tesseractWorker = await Tesseract.createWorker(this.state.language, 1, {
        logger: (m) => {
          if (m.status === 'recognizing text') {
            const pct = Math.round(m.progress * 100);
            this.elements.progressDetail.textContent = `OCR 辨識中... ${pct}%`;
          }
        }
      });
    }

    this.updateProgress(
      35 + (index / total) * 35,
      `處理第 ${index + 1}/${total} 頁`,
      'OCR 辨識中'
    );

    // 執行 OCR
    const { data } = await this.state.tesseractWorker.recognize(page.ocrImageData);

    // 轉換 OCR 結果為文字行
    if (data.lines && data.lines.length > 0) {
      // OCR 圖片的尺寸（scale 1.5）
      const ocrScale = 1.5;
      const ocrWidth = page.width * ocrScale;
      const ocrHeight = page.height * ocrScale;

      return data.lines.map(line => ({
        text: line.text,
        x: (line.bbox.x0 / ocrWidth) * page.width,
        y: (line.bbox.y0 / ocrHeight) * page.height,
        width: ((line.bbox.x1 - line.bbox.x0) / ocrWidth) * page.width,
        height: ((line.bbox.y1 - line.bbox.y0) / ocrHeight) * page.height,
        confidence: line.confidence
      })).filter(line => line.text.trim() && line.confidence > 30);
    }

    return [];
  },

  /**
   * 清理 OCR 資源
   */
  async cleanupOCR() {
    if (this.state.tesseractWorker) {
      await this.state.tesseractWorker.terminate();
      this.state.tesseractWorker = null;
    }
  },

  /**
   * 更新 UI 狀態
   */
  updateUI(status) {
    const { convertBtn, resetBtn, progressSection, successBox } = this.elements;

    switch (status) {
      case 'converting':
        convertBtn.disabled = true;
        convertBtn.innerHTML = '<span class="spinner"></span> 轉換中...';
        progressSection.classList.add('show');
        break;

      case 'success':
        progressSection.classList.remove('show');
        successBox.classList.add('show');
        convertBtn.style.display = 'none';
        resetBtn.style.display = 'block';
        break;

      case 'error':
        progressSection.classList.remove('show');
        convertBtn.innerHTML = '🚀 重新嘗試';
        convertBtn.disabled = false;
        break;
    }
  },

  /**
   * 更新進度
   */
  updateProgress(percent, text, detail) {
    const { progressFill, progressText, progressDetail } = this.elements;
    progressFill.style.width = percent + '%';
    progressText.textContent = text;
    progressDetail.textContent = detail || '';
  },

  /**
   * 顯示錯誤
   */
  showError(message) {
    const { errorBox } = this.elements;
    errorBox.textContent = '❌ ' + message;
    errorBox.classList.add('show');
  },

  /**
   * 隱藏錯誤
   */
  hideError() {
    this.elements.errorBox.classList.remove('show');
  },

  /**
   * 重置應用程式
   */
  reset() {
    const { 
      fileInput, uploadZone, uploadPrompt, fileInfo, 
      convertBtn, resetBtn, successBox, progressSection,
      previewSection, previewContainer, progressFill
    } = this.elements;

    this.state.file = null;
    fileInput.value = '';
    uploadPrompt.style.display = 'block';
    fileInfo.style.display = 'none';
    fileInfo.classList.remove('show');
    uploadZone.classList.remove('has-file');
    successBox.classList.remove('show');
    previewSection.classList.remove('show');
    previewContainer.innerHTML = '';
    convertBtn.style.display = 'block';
    convertBtn.innerHTML = '🚀 開始轉換';
    convertBtn.disabled = true;
    resetBtn.style.display = 'none';
    progressSection.classList.remove('show');
    progressFill.style.width = '0%';
    this.hideError();
  },

  /**
   * 格式化檔案大小
   */
  formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }
};

// 頁面載入完成後初始化
document.addEventListener('DOMContentLoaded', () => {
  App.init();
});
