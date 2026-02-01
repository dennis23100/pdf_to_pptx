/**
 * PDF 轉 PPTX 專業版 - 核心應用程式
 * 
 * 技術棧：
 * - PDF.js: PDF 解析和渲染
 * - Tesseract.js: OCR 文字識別
 * - LaMa ONNX: AI 文字移除（瀏覽器端）
 * - PptxGenJS: PPTX 生成
 * 
 * 作者：Claude AI
 * 授權：MIT License
 */

// ============================================================
// 全域變數和設定
// ============================================================

const CONFIG = {
    // LaMa 模型設定 - 使用較小的模型版本
    LAMA_MODEL_URL: 'https://huggingface.co/Carve/LaMa-ONNX/resolve/main/lama_fp32.onnx',
    // 備選 URL（如果主 URL 失敗）
    LAMA_MODEL_URL_BACKUP: 'https://cdn.jsdelivr.net/gh/nicktomlin/nicktomlin.github.io@main/model/lama_fp32.onnx',
    LAMA_MODEL_SIZE: 50, // MB (實際約 50MB 壓縮後)
    LAMA_INPUT_SIZE: 512,
    
    // 處理設定
    MAX_FILE_SIZE: 100 * 1024 * 1024, // 100MB
    RENDER_SCALE: 2.0, // PDF 渲染比例
    THUMB_SCALE: 0.5, // 縮圖比例
    
    // PPTX 設定
    SLIDE_WIDTH: 13.333, // 英吋 (16:9)
    SLIDE_HEIGHT: 7.5,
    
    // IndexedDB 設定
    DB_NAME: 'PDFtoPPTX_Cache',
    DB_VERSION: 1,
    MODEL_STORE: 'models'
};

// 應用狀態
const state = {
    file: null,
    pages: [],
    selectedPages: new Set(),
    lamaSession: null,
    lamaLoaded: false,
    tesseractWorker: null,
    processing: false,
    pptxBlob: null
};

// ============================================================
// 初始化
// ============================================================

document.addEventListener('DOMContentLoaded', () => {
    initPdfJs();
    initEventListeners();
    initModelCache();
});

function initPdfJs() {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

function initEventListeners() {
    // 上傳區域
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    
    uploadArea.addEventListener('click', () => fileInput.click());
    
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });
    
    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });
    
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleFile(e.dataTransfer.files[0]);
        }
    });
    
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });
    
    // 檔案移除
    document.getElementById('fileRemove').addEventListener('click', resetFile);
    
    // 導航按鈕
    document.getElementById('nextBtn').addEventListener('click', () => goToStep(2));
    document.getElementById('backBtn1').addEventListener('click', () => goToStep(1));
    document.getElementById('startBtn').addEventListener('click', startProcessing);
    document.getElementById('selectAllBtn').addEventListener('click', toggleSelectAll);
    document.getElementById('downloadAgainBtn').addEventListener('click', downloadPptx);
    document.getElementById('restartBtn').addEventListener('click', restart);
    
    // 模式選擇變更
    document.getElementById('modeSelect').addEventListener('change', handleModeChange);
}

// ============================================================
// 模型緩存（IndexedDB）
// ============================================================

let db = null;

async function initModelCache() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(CONFIG.DB_NAME, CONFIG.DB_VERSION);
        
        request.onerror = () => reject(request.error);
        
        request.onsuccess = () => {
            db = request.result;
            resolve();
        };
        
        request.onupgradeneeded = (event) => {
            const database = event.target.result;
            if (!database.objectStoreNames.contains(CONFIG.MODEL_STORE)) {
                database.createObjectStore(CONFIG.MODEL_STORE);
            }
        };
    });
}

async function getCachedModel() {
    if (!db) return null;
    
    return new Promise((resolve, reject) => {
        const transaction = db.transaction([CONFIG.MODEL_STORE], 'readonly');
        const store = transaction.objectStore(CONFIG.MODEL_STORE);
        const request = store.get('lama_model');
        
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function cacheModel(data) {
    if (!db) return;
    
    return new Promise((resolve, reject) => {
        const transaction = db.transaction([CONFIG.MODEL_STORE], 'readwrite');
        const store = transaction.objectStore(CONFIG.MODEL_STORE);
        const request = store.put(data, 'lama_model');
        
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
    });
}

// ============================================================
// 檔案處理
// ============================================================

async function handleFile(file) {
    // 驗證檔案
    const validTypes = ['application/pdf', 'image/png', 'image/jpeg'];
    const isValid = validTypes.includes(file.type) || /\.(pdf|png|jpe?g)$/i.test(file.name);
    
    if (!isValid) {
        showAlert('error', '請選擇 PDF 或圖片檔案（PNG、JPG）');
        return;
    }
    
    if (file.size > CONFIG.MAX_FILE_SIZE) {
        showAlert('error', '檔案過大，請選擇 100MB 以下的檔案');
        return;
    }
    
    state.file = file;
    
    // 顯示檔案資訊
    document.getElementById('fileIcon').textContent = file.type === 'application/pdf' ? '📑' : '🖼️';
    document.getElementById('fileName').textContent = file.name;
    document.getElementById('fileMeta').textContent = `${formatFileSize(file.size)} · ${file.type || '未知類型'}`;
    document.getElementById('fileInfo').classList.add('show');
    
    // 解析檔案
    try {
        showAlert('info', '正在解析檔案...');
        await parseFile(file);
        hideAlert();
        document.getElementById('nextBtn').disabled = false;
        
        // 如果是 AI 模式，預載模型
        const mode = document.getElementById('modeSelect').value;
        if (mode === 'ai') {
            preloadLamaModel();
        }
    } catch (error) {
        console.error('檔案解析失敗:', error);
        showAlert('error', '檔案解析失敗: ' + error.message);
    }
}

async function parseFile(file) {
    state.pages = [];
    
    if (file.type === 'application/pdf') {
        await parsePdf(file);
    } else {
        await parseImage(file);
    }
}

async function parsePdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        
        // 渲染高解析度圖片
        const viewport = page.getViewport({ scale: CONFIG.RENDER_SCALE });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');
        
        await page.render({ canvasContext: ctx, viewport }).promise;
        
        // 生成縮圖
        const thumbCanvas = document.createElement('canvas');
        const thumbWidth = viewport.width * CONFIG.THUMB_SCALE / CONFIG.RENDER_SCALE;
        const thumbHeight = viewport.height * CONFIG.THUMB_SCALE / CONFIG.RENDER_SCALE;
        thumbCanvas.width = thumbWidth;
        thumbCanvas.height = thumbHeight;
        const thumbCtx = thumbCanvas.getContext('2d');
        thumbCtx.drawImage(canvas, 0, 0, thumbWidth, thumbHeight);
        
        state.pages.push({
            pageNum: i,
            canvas: canvas,
            thumb: thumbCanvas.toDataURL('image/jpeg', 0.7),
            width: viewport.width,
            height: viewport.height
        });
    }
}

async function parseImage(file) {
    const dataUrl = await readFileAsDataURL(file);
    
    const img = new Image();
    img.src = dataUrl;
    await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
    });
    
    // 創建 canvas
    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0);
    
    // 生成縮圖
    const thumbCanvas = document.createElement('canvas');
    const scale = Math.min(300 / img.width, 200 / img.height);
    thumbCanvas.width = img.width * scale;
    thumbCanvas.height = img.height * scale;
    const thumbCtx = thumbCanvas.getContext('2d');
    thumbCtx.drawImage(img, 0, 0, thumbCanvas.width, thumbCanvas.height);
    
    state.pages.push({
        pageNum: 1,
        canvas: canvas,
        thumb: thumbCanvas.toDataURL('image/jpeg', 0.7),
        width: img.width,
        height: img.height
    });
}

function resetFile() {
    state.file = null;
    state.pages = [];
    state.selectedPages.clear();
    
    document.getElementById('fileInfo').classList.remove('show');
    document.getElementById('fileInput').value = '';
    document.getElementById('nextBtn').disabled = true;
    hideAlert();
}

// ============================================================
// 步驟導航
// ============================================================

function goToStep(stepNum) {
    // 更新步驟指示器
    document.querySelectorAll('.step-item').forEach((item, index) => {
        item.classList.remove('active', 'completed');
        if (index + 1 < stepNum) {
            item.classList.add('completed');
        } else if (index + 1 === stepNum) {
            item.classList.add('active');
        }
    });
    
    // 隱藏所有內容
    document.querySelectorAll('.step-content').forEach(el => {
        el.style.display = 'none';
        el.classList.remove('show');
    });
    
    // 顯示當前步驟
    const stepEl = document.getElementById(`step${stepNum}`);
    stepEl.style.display = 'block';
    
    if (stepNum === 2) {
        renderPreviewGrid();
    } else if (stepNum === 3) {
        stepEl.classList.add('show');
    } else if (stepNum === 4) {
        stepEl.classList.add('show');
    }
}

function renderPreviewGrid() {
    const grid = document.getElementById('previewGrid');
    grid.innerHTML = '';
    
    // 預設全選
    state.selectedPages.clear();
    
    state.pages.forEach((page, index) => {
        state.selectedPages.add(index);
        
        const item = document.createElement('div');
        item.className = 'preview-item selected';
        item.dataset.index = index;
        item.innerHTML = `
            <img src="${page.thumb}" alt="第 ${page.pageNum} 頁">
            <div class="page-num">第 ${page.pageNum} 頁</div>
            <div class="check-mark">✓</div>
        `;
        
        item.addEventListener('click', () => togglePageSelection(index, item));
        grid.appendChild(item);
    });
    
    updateSelectedCount();
}

function togglePageSelection(index, element) {
    if (state.selectedPages.has(index)) {
        state.selectedPages.delete(index);
        element.classList.remove('selected');
    } else {
        state.selectedPages.add(index);
        element.classList.add('selected');
    }
    updateSelectedCount();
}

function toggleSelectAll() {
    const allSelected = state.selectedPages.size === state.pages.length;
    
    document.querySelectorAll('.preview-item').forEach((item, index) => {
        if (allSelected) {
            state.selectedPages.delete(index);
            item.classList.remove('selected');
        } else {
            state.selectedPages.add(index);
            item.classList.add('selected');
        }
    });
    
    updateSelectedCount();
}

function updateSelectedCount() {
    document.getElementById('selectedCount').textContent = state.selectedPages.size;
    document.getElementById('startBtn').disabled = state.selectedPages.size === 0;
}

// ============================================================
// 模式處理
// ============================================================

function handleModeChange() {
    const mode = document.getElementById('modeSelect').value;
    const modelStatus = document.getElementById('modelStatus');
    
    if (mode === 'ai') {
        modelStatus.classList.add('show');
        if (state.file) {
            preloadLamaModel();
        }
    } else {
        modelStatus.classList.remove('show');
    }
}

// ============================================================
// LaMa 模型載入
// ============================================================

async function preloadLamaModel() {
    if (state.lamaLoaded) {
        updateModelStatus('ready', '✅', 'AI 模型已就緒');
        return;
    }
    
    const modelStatus = document.getElementById('modelStatus');
    modelStatus.classList.add('show');
    updateModelStatus('loading', '⏳', 'AI 模型載入中...');
    
    try {
        // 檢查緩存
        let modelData = await getCachedModel();
        
        if (modelData) {
            updateModelStatus('loading', '📦', '從緩存載入模型...');
            document.getElementById('modelProgressFill').style.width = '50%';
        } else {
            updateModelStatus('loading', '📥', `下載 AI 模型中（約 ${CONFIG.LAMA_MODEL_SIZE}MB）...`);
            document.getElementById('modelStatusDetail').textContent = 
                '首次使用需下載，下載後會自動緩存到瀏覽器（可能需要 1-3 分鐘）';
            
            // 嘗試下載模型
            try {
                modelData = await downloadModelWithProgress(CONFIG.LAMA_MODEL_URL, (progress) => {
                    document.getElementById('modelProgressFill').style.width = `${progress * 50}%`;
                });
            } catch (downloadError) {
                console.warn('主 URL 下載失敗，嘗試備選 URL:', downloadError);
                // 如果主 URL 失敗，嘗試備選 URL
                updateModelStatus('loading', '🔄', '切換備選下載源...');
                try {
                    modelData = await downloadModelWithProgress(CONFIG.LAMA_MODEL_URL_BACKUP, (progress) => {
                        document.getElementById('modelProgressFill').style.width = `${progress * 50}%`;
                    });
                } catch (backupError) {
                    throw new Error('模型下載失敗，請檢查網路連線或稍後重試');
                }
            }
            
            // 緩存模型
            try {
                await cacheModel(modelData);
            } catch (cacheError) {
                console.warn('模型緩存失敗:', cacheError);
                // 緩存失敗不影響使用
            }
        }
        
        // 初始化 ONNX Session
        updateModelStatus('loading', '🔧', '初始化 AI 引擎...');
        document.getElementById('modelProgressFill').style.width = '75%';
        
        state.lamaSession = await ort.InferenceSession.create(modelData, {
            executionProviders: ['wasm'],
            graphOptimizationLevel: 'all'
        });
        
        state.lamaLoaded = true;
        document.getElementById('modelProgressFill').style.width = '100%';
        updateModelStatus('ready', '✅', 'AI 模型已就緒');
        
    } catch (error) {
        console.error('模型載入失敗:', error);
        updateModelStatus('error', '❌', 'AI 模型載入失敗');
        document.getElementById('modelStatusDetail').textContent = 
            error.message + '。建議使用「背景色覆蓋」模式作為替代方案。';
        
        // 自動切換到備選模式
        document.getElementById('modeSelect').value = 'overlay';
        showAlert('warning', 'AI 模型載入失敗，已自動切換到「背景色覆蓋」模式');
    }
}

async function downloadModelWithProgress(url, onProgress) {
    const response = await fetch(url);
    
    if (!response.ok) {
        throw new Error(`下載失敗: ${response.status}`);
    }
    
    const contentLength = response.headers.get('content-length');
    const total = parseInt(contentLength, 10);
    let loaded = 0;
    
    const reader = response.body.getReader();
    const chunks = [];
    
    while (true) {
        const { done, value } = await reader.read();
        
        if (done) break;
        
        chunks.push(value);
        loaded += value.length;
        
        if (total) {
            onProgress(loaded / total);
        }
    }
    
    // 合併所有 chunks
    const allChunks = new Uint8Array(loaded);
    let position = 0;
    for (const chunk of chunks) {
        allChunks.set(chunk, position);
        position += chunk.length;
    }
    
    return allChunks.buffer;
}

function updateModelStatus(type, icon, text) {
    const modelStatus = document.getElementById('modelStatus');
    document.getElementById('modelStatusIcon').textContent = icon;
    document.getElementById('modelStatusText').textContent = text;
    
    modelStatus.classList.remove('ready');
    if (type === 'ready') {
        modelStatus.classList.add('ready');
    }
}

// ============================================================
// 主要處理流程
// ============================================================

async function startProcessing() {
    if (state.processing) return;
    state.processing = true;
    
    const mode = document.getElementById('modeSelect').value;
    const lang = document.getElementById('langSelect').value;
    const quality = document.getElementById('qualitySelect').value;
    
    goToStep(3);
    
    try {
        // 步驟 1: 載入 AI 模型（如果需要）
        updateProcessingStep(1, 'active');
        updateProcessingStatus('載入 AI 模型...', '準備處理環境');
        updateProcessingProgress(5);
        
        if (mode === 'ai' && !state.lamaLoaded) {
            await preloadLamaModel();
        }
        
        // 初始化 Tesseract (修正版：增加錯誤處理與降級機制)
        if (mode !== 'image') {
            try {
                // 嘗試初始化 Worker
                state.tesseractWorker = await Tesseract.createWorker(lang, 1, {
                    logger: m => {
                        if (m.status === 'recognizing text') {
                            updateProcessingProgress(20 + m.progress * 20);
                        }
                    },
                    // 強制使用指定的 Worker 路徑，避免自動抓取錯誤版本
                    workerPath: 'https://unpkg.com/tesseract.js@v4.1.1/dist/worker.min.js',
                    corePath: 'https://unpkg.com/tesseract.js-core@v4.0.4/tesseract-core.wasm.js'
                });
            } catch (err) {
                console.warn('OCR 初始化失敗，嘗試單線程模式', err);
                // 備用方案：如果不支援 Worker，嘗試不帶參數初始化（依賴 CDN 預設）
                 state.tesseractWorker = await Tesseract.createWorker(lang);
            }
        }
        
        updateProcessingStep(1, 'completed');
        
        // 處理每一頁
        const selectedIndices = Array.from(state.selectedPages).sort((a, b) => a - b);
        const pptx = new PptxGenJS();
        pptx.layout = 'LAYOUT_WIDE';
        
        for (let i = 0; i < selectedIndices.length; i++) {
            const pageIndex = selectedIndices[i];
            const page = state.pages[pageIndex];
            const progress = 20 + (i / selectedIndices.length) * 60;
            
            updateProcessingStatus(
                `處理第 ${i + 1}/${selectedIndices.length} 頁`,
                `第 ${page.pageNum} 頁`
            );
            
            // 步驟 2: OCR
            updateProcessingStep(2, 'active');
            updateProcessingProgress(progress);
            
            let textData = null;
            if (mode !== 'image') {
                textData = await performOCR(page.canvas, lang);
            }
            
            updateProcessingStep(2, 'completed');
            
            // 步驟 3: AI 處理
            updateProcessingStep(3, 'active');
            
            let backgroundCanvas = page.canvas;
            if (mode === 'ai' && textData && textData.lines.length > 0) {
                backgroundCanvas = await removeTextWithLama(page.canvas, textData);
            }
            
            updateProcessingStep(3, 'completed');
            
            // 步驟 4: 生成 PPTX
            updateProcessingStep(4, 'active');
            
            await createSlide(pptx, backgroundCanvas, textData, mode);
        }
        
        // 清理 Tesseract
        if (state.tesseractWorker) {
            await state.tesseractWorker.terminate();
            state.tesseractWorker = null;
        }
        
        // 生成並下載
        updateProcessingStatus('生成 PPTX 檔案...', '即將完成');
        updateProcessingProgress(95);
        
        const fileName = state.file.name.replace(/\.[^.]+$/, '') + '_editable.pptx';
        state.pptxBlob = await pptx.write({ outputType: 'blob' });
        
        // 下載
        downloadBlob(state.pptxBlob, fileName);
        
        updateProcessingStep(4, 'completed');
        updateProcessingProgress(100);
        
        // 完成
        setTimeout(() => {
            document.getElementById('resultDetail').textContent = 
                `已成功處理 ${selectedIndices.length} 頁，檔案已自動下載`;
            goToStep(4);
        }, 500);
        
    } catch (error) {
        console.error('處理失敗:', error);
        showAlert('error', '處理失敗: ' + error.message);
        goToStep(1);
    } finally {
        state.processing = false;
    }
}

// ============================================================
// OCR 處理
// ============================================================

async function performOCR(canvas, lang) {
    if (!state.tesseractWorker) return { text: '', lines: [], words: [] };

    try {
        const { data } = await state.tesseractWorker.recognize(canvas);
        
        return {
            text: data.text,
            lines: data.lines.filter(line => line.confidence > 30 && line.text.trim()),
            words: data.words
        };
    } catch (error) {
        console.warn('OCR 辨識錯誤，跳過此頁:', error);
        return { text: '', lines: [], words: [] };
    }
}

// ============================================================
// LaMa AI 文字移除
// ============================================================

async function removeTextWithLama(canvas, textData) {
    if (!state.lamaSession) {
        console.warn('LaMa 模型未載入，跳過 AI 處理');
        return canvas;
    }
    
    const width = canvas.width;
    const height = canvas.height;
    
    // 創建 mask（標記文字區域）
    const maskCanvas = document.createElement('canvas');
    maskCanvas.width = width;
    maskCanvas.height = height;
    const maskCtx = maskCanvas.getContext('2d');
    maskCtx.fillStyle = 'black';
    maskCtx.fillRect(0, 0, width, height);
    
    // 在 mask 上標記文字區域（白色）
    maskCtx.fillStyle = 'white';
    for (const line of textData.lines) {
        const bbox = line.bbox;
        if (bbox) {
            // 稍微擴大區域以確保完全覆蓋
            const padding = 5;
            maskCtx.fillRect(
                bbox.x0 - padding,
                bbox.y0 - padding,
                bbox.x1 - bbox.x0 + padding * 2,
                bbox.y1 - bbox.y0 + padding * 2
            );
        }
    }
    
    // 分塊處理（因為 LaMa 只支援 512x512）
    const resultCanvas = document.createElement('canvas');
    resultCanvas.width = width;
    resultCanvas.height = height;
    const resultCtx = resultCanvas.getContext('2d');
    
    // 先複製原圖
    resultCtx.drawImage(canvas, 0, 0);
    
    // 計算需要處理的塊
    const blockSize = CONFIG.LAMA_INPUT_SIZE;
    const overlap = 64; // 重疊區域以避免接縫
    
    for (let y = 0; y < height; y += blockSize - overlap) {
        for (let x = 0; x < width; x += blockSize - overlap) {
            // 檢查這個區域是否有 mask
            const blockMaskData = maskCtx.getImageData(x, y, 
                Math.min(blockSize, width - x), 
                Math.min(blockSize, height - y)
            );
            
            // 檢查是否有白色像素（需要處理的區域）
            let hasMask = false;
            for (let i = 0; i < blockMaskData.data.length; i += 4) {
                if (blockMaskData.data[i] > 128) {
                    hasMask = true;
                    break;
                }
            }
            
            if (!hasMask) continue;
            
            // 提取區塊
            const blockCanvas = document.createElement('canvas');
            blockCanvas.width = blockSize;
            blockCanvas.height = blockSize;
            const blockCtx = blockCanvas.getContext('2d');
            
            // 複製圖像區塊
            blockCtx.drawImage(canvas, -x, -y);
            
            // 複製 mask 區塊
            const blockMaskCanvas = document.createElement('canvas');
            blockMaskCanvas.width = blockSize;
            blockMaskCanvas.height = blockSize;
            const blockMaskCtx = blockMaskCanvas.getContext('2d');
            blockMaskCtx.drawImage(maskCanvas, -x, -y);
            
            try {
                // 執行 LaMa 推理
                const inpaintedBlock = await runLamaInference(blockCanvas, blockMaskCanvas);
                
                // 將結果繪製回結果 canvas
                resultCtx.drawImage(inpaintedBlock, x, y);
            } catch (error) {
                console.warn('LaMa 推理失敗，跳過此區塊:', error);
            }
        }
    }
    
    return resultCanvas;
}

async function runLamaInference(imageCanvas, maskCanvas) {
    const size = CONFIG.LAMA_INPUT_SIZE;
    
    // 確保 canvas 是 512x512
    const resizedImage = resizeCanvas(imageCanvas, size, size);
    const resizedMask = resizeCanvas(maskCanvas, size, size);
    
    // 轉換為 tensor
    const imageTensor = canvasToTensor(resizedImage);
    const maskTensor = canvasToMaskTensor(resizedMask);
    
    // 執行推理
    const feeds = {
        image: imageTensor,
        mask: maskTensor
    };
    
    const results = await state.lamaSession.run(feeds);
    const output = results[Object.keys(results)[0]];
    
    // 轉換輸出為 canvas
    const outputCanvas = tensorToCanvas(output, size, size);
    
    return outputCanvas;
}

function resizeCanvas(canvas, width, height) {
    const resized = document.createElement('canvas');
    resized.width = width;
    resized.height = height;
    const ctx = resized.getContext('2d');
    ctx.drawImage(canvas, 0, 0, width, height);
    return resized;
}

function canvasToTensor(canvas) {
    const ctx = canvas.getContext('2d');
    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const data = imageData.data;
    
    const float32Data = new Float32Array(3 * canvas.width * canvas.height);
    
    for (let i = 0; i < canvas.width * canvas.height; i++) {
        float32Data[i] = data[i * 4] / 255.0; // R
        float32Data[i + canvas.width * canvas.height] = data[i * 4 + 1] / 255.0; // G
        float32Data[i + 2 * canvas.width * canvas.height] = data[i * 4 + 2] / 255.0; // B
    }
    
    return new ort.Tensor('float32', float32Data, [1, 3, canvas.height, canvas.width]);
}

function canvasToMaskTensor(canvas) {
    const ctx = canvas.getContext('2d');
    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const data = imageData.data;
    
    const float32Data = new Float32Array(canvas.width * canvas.height);
    
    for (let i = 0; i < canvas.width * canvas.height; i++) {
        // 二值化 mask
        float32Data[i] = data[i * 4] > 128 ? 1.0 : 0.0;
    }
    
    return new ort.Tensor('float32', float32Data, [1, 1, canvas.height, canvas.width]);
}

function tensorToCanvas(tensor, width, height) {
    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    
    const imageData = ctx.createImageData(width, height);
    const data = tensor.data;
    
    for (let i = 0; i < width * height; i++) {
        imageData.data[i * 4] = Math.max(0, Math.min(255, data[i] * 255)); // R
        imageData.data[i * 4 + 1] = Math.max(0, Math.min(255, data[i + width * height] * 255)); // G
        imageData.data[i * 4 + 2] = Math.max(0, Math.min(255, data[i + 2 * width * height] * 255)); // B
        imageData.data[i * 4 + 3] = 255; // A
    }
    
    ctx.putImageData(imageData, 0, 0);
    return canvas;
}

// ============================================================
// PPTX 生成
// ============================================================

async function createSlide(pptx, backgroundCanvas, textData, mode) {
    const slide = pptx.addSlide();
    
    // 添加背景圖
    const bgDataUrl = backgroundCanvas.toDataURL('image/jpeg', 0.9);
    slide.addImage({
        data: bgDataUrl,
        x: 0, y: 0,
        w: '100%', h: '100%'
    });
    
    // 如果不是純圖片模式，添加可編輯文字
    if (mode !== 'image' && textData && textData.lines) {
        const imgWidth = backgroundCanvas.width;
        const imgHeight = backgroundCanvas.height;
        
        for (const line of textData.lines) {
            if (!line.text || !line.bbox) continue;
            
            const bbox = line.bbox;
            
            // 座標轉換
            const x = (bbox.x0 / imgWidth) * CONFIG.SLIDE_WIDTH;
            const y = (bbox.y0 / imgHeight) * CONFIG.SLIDE_HEIGHT;
            const w = Math.max(0.5, ((bbox.x1 - bbox.x0) / imgWidth) * CONFIG.SLIDE_WIDTH);
            const h = Math.max(0.3, ((bbox.y1 - bbox.y0) / imgHeight) * CONFIG.SLIDE_HEIGHT);
            
            // 計算字體大小
            const fontSize = Math.max(8, Math.min(36, (bbox.y1 - bbox.y0) * 0.5));
            
            // 如果是覆蓋模式，添加背景色矩形
            if (mode === 'overlay') {
                slide.addShape('rect', {
                    x: x,
                    y: y,
                    w: w * 1.05,
                    h: h * 1.1,
                    fill: { color: 'F5F0E8' },
                    line: { color: 'F5F0E8', width: 0 }
                });
            }
            
            // 添加文字
            slide.addText(line.text, {
                x: x,
                y: y,
                w: w * 1.1,
                h: h * 1.3,
                fontSize: fontSize,
                fontFace: 'Microsoft JhengHei',
                color: '333333',
                valign: 'top'
            });
        }
    }
}

// ============================================================
// 輔助函數
// ============================================================

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function readFileAsDataURL(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error('檔案讀取失敗'));
        reader.readAsDataURL(file);
    });
}

function showAlert(type, message) {
    const alertEl = document.getElementById('alertInfo');
    alertEl.className = `alert alert-${type} show`;
    document.getElementById('alertInfoText').textContent = message;
}

function hideAlert() {
    document.getElementById('alertInfo').classList.remove('show');
}

function updateProcessingStatus(status, detail) {
    document.getElementById('processingStatus').textContent = status;
    document.getElementById('processingDetail').textContent = detail;
}

function updateProcessingProgress(percent) {
    document.getElementById('processingProgressFill').style.width = `${percent}%`;
}

function updateProcessingStep(stepNum, status) {
    const stepEl = document.getElementById(`pStep${stepNum}`);
    stepEl.classList.remove('active', 'completed');
    if (status) {
        stepEl.classList.add(status);
    }
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function downloadPptx() {
    if (state.pptxBlob) {
        const fileName = state.file.name.replace(/\.[^.]+$/, '') + '_editable.pptx';
        downloadBlob(state.pptxBlob, fileName);
    }
}

function restart() {
    state.file = null;
    state.pages = [];
    state.selectedPages.clear();
    state.pptxBlob = null;
    
    document.getElementById('fileInfo').classList.remove('show');
    document.getElementById('fileInput').value = '';
    document.getElementById('nextBtn').disabled = true;
    
    goToStep(1);
}
