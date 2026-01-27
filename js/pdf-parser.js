/**
 * PDF Parser Module
 * 負責解析 PDF，提取文字、位置和頁面圖片
 */

const PDFParser = {
  /**
   * 解析 PDF 檔案
   * @param {ArrayBuffer} arrayBuffer - PDF 檔案的 ArrayBuffer
   * @param {Function} onProgress - 進度回調
   * @returns {Object} 解析結果
   */
  async parse(arrayBuffer, onProgress = () => {}) {
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const numPages = pdf.numPages;
    const pages = [];

    for (let i = 1; i <= numPages; i++) {
      onProgress({
        current: i,
        total: numPages,
        phase: 'parsing',
        message: `解析第 ${i}/${numPages} 頁...`
      });

      const pageData = await this.parsePage(pdf, i);
      pages.push(pageData);
    }

    return {
      numPages,
      pages
    };
  },

  /**
   * 解析單一頁面
   * @param {PDFDocumentProxy} pdf - PDF 文件物件
   * @param {number} pageNum - 頁碼
   * @returns {Object} 頁面資料
   */
  async parsePage(pdf, pageNum) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1 });

    // 提取原生文字層
    const textContent = await page.getTextContent();
    const textItems = this.extractTextItems(textContent, viewport);

    // 渲染頁面為圖片（高解析度）
    const imageData = await this.renderPageToImage(page, 2.5);

    // 渲染較小圖片用於 OCR（如果需要）
    const ocrImageData = await this.renderPageToImage(page, 1.5);

    return {
      pageNum,
      width: viewport.width,
      height: viewport.height,
      textItems,
      hasNativeText: textItems.length > 0,
      imageData,
      ocrImageData
    };
  },

  /**
   * 提取文字項目及其位置
   * @param {Object} textContent - PDF.js 的 textContent 物件
   * @param {Object} viewport - 頁面視窗
   * @returns {Array} 文字項目陣列
   */
  extractTextItems(textContent, viewport) {
    if (!textContent || !textContent.items) return [];

    return textContent.items
      .filter(item => item.str && item.str.trim())
      .map(item => {
        const tx = item.transform;
        
        // transform 格式: [scaleX, skewX, skewY, scaleY, x, y]
        // PDF 座標系統: 原點在左下角，Y 軸向上
        const x = tx[4];
        const y = viewport.height - tx[5]; // 轉換 Y 軸（從下到上 → 從上到下）
        const fontSize = Math.abs(tx[0]) || Math.abs(tx[3]) || 12;
        const width = item.width || (item.str.length * fontSize * 0.6);
        const height = item.height || fontSize;

        return {
          text: item.str,
          x: x,
          y: y - height, // 調整到左上角
          width: width,
          height: height,
          fontSize: fontSize,
          fontName: item.fontName || 'default'
        };
      });
  },

  /**
   * 將頁面渲染為圖片
   * @param {PDFPageProxy} page - PDF 頁面物件
   * @param {number} scale - 縮放比例
   * @returns {string} Base64 圖片資料
   */
  async renderPageToImage(page, scale = 2) {
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');

    await page.render({
      canvasContext: ctx,
      viewport: viewport
    }).promise;

    return canvas.toDataURL('image/jpeg', 0.92);
  },

  /**
   * 將文字項目按行分組
   * @param {Array} textItems - 文字項目陣列
   * @param {number} threshold - Y 座標閾值
   * @returns {Array} 分組後的行陣列
   */
  groupTextByLines(textItems, threshold = 10) {
    if (!textItems || textItems.length === 0) return [];

    // 按 Y 座標排序
    const sorted = [...textItems].sort((a, b) => a.y - b.y || a.x - b.x);
    
    const lines = [];
    let currentLine = [sorted[0]];

    for (let i = 1; i < sorted.length; i++) {
      const item = sorted[i];
      const prevItem = currentLine[0];

      // 如果 Y 座標接近，視為同一行
      if (Math.abs(item.y - prevItem.y) < threshold) {
        currentLine.push(item);
      } else {
        // 新的一行
        lines.push(this.mergeLine(currentLine));
        currentLine = [item];
      }
    }

    // 處理最後一行
    if (currentLine.length > 0) {
      lines.push(this.mergeLine(currentLine));
    }

    return lines;
  },

  /**
   * 合併同一行的文字項目
   * @param {Array} lineItems - 同一行的文字項目
   * @returns {Object} 合併後的行物件
   */
  mergeLine(lineItems) {
    // 按 X 座標排序
    const sorted = lineItems.sort((a, b) => a.x - b.x);
    
    // 合併文字
    const text = sorted.map(item => item.text).join(' ');
    
    // 計算邊界框
    const x = Math.min(...sorted.map(item => item.x));
    const y = Math.min(...sorted.map(item => item.y));
    const maxX = Math.max(...sorted.map(item => item.x + item.width));
    const maxY = Math.max(...sorted.map(item => item.y + item.height));
    
    // 平均字體大小
    const avgFontSize = sorted.reduce((sum, item) => sum + item.fontSize, 0) / sorted.length;

    return {
      text: text.trim(),
      x: x,
      y: y,
      width: maxX - x,
      height: maxY - y,
      fontSize: avgFontSize
    };
  }
};

// 如果是模組環境則導出
if (typeof module !== 'undefined' && module.exports) {
  module.exports = PDFParser;
}
