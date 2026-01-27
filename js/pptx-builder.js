/**
 * PPTX Builder Module
 * 負責建構 PowerPoint 檔案
 */

const PPTXBuilder = {
  // 投影片尺寸（英寸）
  SLIDE_WIDTH: 10,      // 16:9 標準寬度
  SLIDE_HEIGHT: 5.625,  // 16:9 標準高度

  /**
   * 建構 PPTX 檔案
   * @param {Array} pages - 頁面資料陣列
   * @param {Object} options - 選項
   * @param {Function} onProgress - 進度回調
   * @returns {PptxGenJS} PPTX 物件
   */
  async build(pages, options = {}, onProgress = () => {}) {
    const pptx = new PptxGenJS();
    
    // 設定投影片佈局
    pptx.layout = 'LAYOUT_16x9';
    
    // 設定屬性
    pptx.author = 'PDF to PPTX Converter';
    pptx.title = options.title || 'Converted Presentation';
    pptx.subject = 'Converted from PDF';

    const totalPages = pages.length;

    for (let i = 0; i < totalPages; i++) {
      const page = pages[i];
      
      onProgress({
        current: i + 1,
        total: totalPages,
        phase: 'building',
        message: `建構投影片 ${i + 1}/${totalPages}...`
      });

      await this.addSlide(pptx, page, options);
    }

    return pptx;
  },

  /**
   * 添加投影片
   * @param {PptxGenJS} pptx - PPTX 物件
   * @param {Object} page - 頁面資料
   * @param {Object} options - 選項
   */
  async addSlide(pptx, page, options = {}) {
    const slide = pptx.addSlide();

    // 1. 添加背景圖片
    slide.addImage({
      data: page.imageData,
      x: 0,
      y: 0,
      w: '100%',
      h: '100%'
    });

    // 2. 添加文字方塊（如果有文字資料）
    const textData = page.textLines || page.ocrLines || [];
    
    if (textData.length > 0 && options.addTextBoxes !== false) {
      this.addTextBoxes(slide, textData, page, options);
    }

    // 3. 添加頁碼（可選）
    if (options.showPageNumber) {
      slide.addText(`${page.pageNum}`, {
        x: this.SLIDE_WIDTH - 0.5,
        y: this.SLIDE_HEIGHT - 0.3,
        w: 0.4,
        h: 0.25,
        fontSize: 8,
        color: 'AAAAAA',
        align: 'right'
      });
    }
  },

  /**
   * 添加文字方塊到投影片
   * @param {Slide} slide - 投影片物件
   * @param {Array} textLines - 文字行陣列
   * @param {Object} page - 頁面資料
   * @param {Object} options - 選項
   */
  addTextBoxes(slide, textLines, page, options = {}) {
    const pageWidth = page.width;
    const pageHeight = page.height;
    
    // 字體設定
    const fontFace = options.fontFace || 'Microsoft JhengHei';
    const textColor = options.textColor || '000000';
    const transparency = options.textTransparency !== undefined ? options.textTransparency : 100;

    textLines.forEach(line => {
      if (!line.text || !line.text.trim()) return;

      // 轉換座標：從 PDF 座標到 PPTX 座標（英寸）
      const x = (line.x / pageWidth) * this.SLIDE_WIDTH;
      const y = (line.y / pageHeight) * this.SLIDE_HEIGHT;
      const w = Math.max((line.width / pageWidth) * this.SLIDE_WIDTH, 0.5);
      const h = Math.max((line.height / pageHeight) * this.SLIDE_HEIGHT, 0.2);

      // 計算字體大小
      // 基於方塊高度估算，並限制範圍
      const estimatedFontSize = Math.round(h * 72 * 0.65);
      const fontSize = Math.max(8, Math.min(estimatedFontSize, 36));

      // 添加文字方塊
      slide.addText(line.text, {
        x: x,
        y: y,
        w: w * 1.15,  // 稍微加寬以確保文字不被截斷
        h: h * 1.2,   // 稍微加高
        fontSize: fontSize,
        fontFace: fontFace,
        color: textColor,
        valign: 'top',
        wrap: true,
        transparency: transparency,
        // 讓文字方塊可以被選取和編輯
        hyperlink: null
      });
    });
  },

  /**
   * 下載 PPTX 檔案
   * @param {PptxGenJS} pptx - PPTX 物件
   * @param {string} fileName - 檔案名稱
   */
  async download(pptx, fileName) {
    await pptx.writeFile({ fileName: fileName });
  },

  /**
   * 獲取 PPTX Blob
   * @param {PptxGenJS} pptx - PPTX 物件
   * @returns {Blob} PPTX Blob
   */
  async getBlob(pptx) {
    return await pptx.write({ outputType: 'blob' });
  }
};

// 如果是模組環境則導出
if (typeof module !== 'undefined' && module.exports) {
  module.exports = PPTXBuilder;
}
