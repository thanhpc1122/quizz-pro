/**
 * ============================================================================
 * NEBULAQUIZ STUDIO — CORE APPLICATION LOGIC (PHASE 3.3)
 * Pure Vanilla JavaScript (ES6+) | Client-Side .DOCX Parser & Quiz Engine
 * ============================================================================
 */

(() => {
  'use strict';

  /* ==========================================================================
     1. WEB AUDIO SYNTHESIZER (HIỆU ỨNG ÂM THANH KHÔNG CẦN FILE NGOÀI)
     ========================================================================== */
  class SoundEffects {
    constructor() {
      this.enabled = true;
      this.ctx = null;
    }

    initContext() {
      if (!this.ctx && (window.AudioContext || window.webkitAudioContext)) {
        const AudioCtx = window.AudioContext || window.webkitAudioContext;
        this.ctx = new AudioCtx();
      }
      if (this.ctx && this.ctx.state === 'suspended') {
        this.ctx.resume();
      }
    }

    playTone(freq, type = 'sine', duration = 0.12, delay = 0, volume = 0.12) {
      if (!this.enabled) return;
      try {
        this.initContext();
        if (!this.ctx) return;

        const osc = this.ctx.createOscillator();
        const gain = this.ctx.createGain();
        const startTime = this.ctx.currentTime + delay;

        osc.type = type;
        osc.frequency.setValueAtTime(freq, startTime);

        gain.gain.setValueAtTime(volume, startTime);
        gain.gain.exponentialRampToValueAtTime(0.001, startTime + duration);

        osc.connect(gain);
        gain.connect(this.ctx.destination);

        osc.start(startTime);
        osc.stop(startTime + duration);
      } catch (_) {
        // Bỏ qua nếu trình duyệt chặn tự động phát âm thanh
      }
    }

    playSelect() {
      this.playTone(520, 'triangle', 0.07, 0, 0.08);
    }

    playCorrect() {
      this.playTone(587.33, 'sine', 0.12, 0, 0.14);     // D5
      this.playTone(880, 'triangle', 0.22, 0.09, 0.16); // A5
    }

    playWrong() {
      this.playTone(240, 'sawtooth', 0.14, 0, 0.11);
      this.playTone(180, 'sawtooth', 0.22, 0.1, 0.11);
    }

    playComplete() {
      const notes = [523.25, 659.25, 783.99, 1046.5]; // C5, E5, G5, C6
      notes.forEach((n, idx) => this.playTone(n, 'triangle', 0.18, idx * 0.09, 0.14));
    }
  }

  const sfx = new SoundEffects();

  /* ==========================================================================
     2. FILE READER & NATIVE .DOCX / .TXT / .JSON EXTRACTOR
     ========================================================================== */
  class DocumentExtractor {
    /**
     * Đọc file người dùng tải lên và trả về chuỗi văn bản thô hoặc dữ liệu JSON
     * @param {File} file
     * @param {Function} onProgress
     */
    static async extractContent(file, onProgress) {
      const fileName = file.name || 'document';
      const ext = fileName.split('.').pop().toLowerCase();

      if (!['docx', 'txt', 'json'].includes(ext)) {
        throw new Error(
          `Định dạng ".${ext}" không được hỗ trợ. Vui lòng chọn file Word (.docx), .txt hoặc .json.`
        );
      }

      if (ext === 'txt' || ext === 'json') {
        onProgress('Đang đọc nội dung văn bản...');
        const text = await file.text();
        return { ext, rawText: text };
      }

      // Xử lý file .docx
      onProgress('Đang đọc cấu trúc file Word (.docx)...');
      const arrayBuffer = await file.arrayBuffer();

      // Ưu tiên 1: Sử dụng thư viện Mammoth.js (nếu đã nạp qua CDN)
      if (window.mammoth && typeof window.mammoth.extractRawText === 'function') {
        try {
          const result = await window.mammoth.extractRawText({ arrayBuffer });
          if (result && typeof result.value === 'string' && result.value.trim().length > 0) {
            return { ext: 'docx', rawText: result.value };
          }
        } catch (mammothErr) {
          console.warn('Mammoth.js gặp lỗi, chuyển sang bộ giải nén XML nội bộ:', mammothErr);
        }
      }

      // Ưu tiên 2: Fallback giải nén trực tiếp word/document.xml từ ZIP container của .docx
      onProgress('Đang trích xuất XML trực tiếp từ gói .docx...');
      const xmlText = await DocumentExtractor.extractDocxXmlFallback(arrayBuffer);
      return { ext: 'docx', rawText: xmlText };
    }

    /**
     * Bộ giải nén .docx thuần JavaScript (sử dụng DecompressionStream chuẩn Web API)
     * Giúp đọc được file .docx ngay cả khi offline không có CDN Mammoth
     */
    static async extractDocxXmlFallback(arrayBuffer) {
      const bytes = new Uint8Array(arrayBuffer);
      const view = new DataView(arrayBuffer);

      // Kiểm tra chữ ký ZIP (PK\x03\x04)
      if (bytes.length < 4 || view.getUint32(0, true) !== 0x04034b50) {
        throw new Error('File .docx không hợp lệ hoặc bị hỏng (không đúng chuẩn OpenXML ZIP).');
      }

      // Tìm End of Central Directory (EOCD)
      let eocdOffset = -1;
      for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
        if (view.getUint32(i, true) === 0x06054b50) {
          eocdOffset = i;
          break;
        }
      }

      if (eocdOffset === -1) {
        throw new Error('Không tìm thấy danh mục Central Directory trong file .docx.');
      }

      const cdEntries = view.getUint16(eocdOffset + 10, true);
      const cdOffset = view.getUint32(eocdOffset + 16, true);
      let ptr = cdOffset;
      const decoder = new TextDecoder('utf-8');

      for (let i = 0; i < cdEntries; i++) {
        if (view.getUint32(ptr, true) !== 0x02014b50) break;

        const compressionMethod = view.getUint16(ptr + 10, true);
        const compressedSize = view.getUint32(ptr + 20, true);
        const fileNameLen = view.getUint16(ptr + 28, true);
        const extraLen = view.getUint16(ptr + 30, true);
        const commentLen = view.getUint16(ptr + 32, true);
        const localHeaderOffset = view.getUint32(ptr + 42, true);

        const entryName = decoder.decode(bytes.subarray(ptr + 46, ptr + 46 + fileNameLen));

        if (entryName === 'word/document.xml') {
          const localNameLen = view.getUint16(localHeaderOffset + 26, true);
          const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
          const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
          const compressedData = bytes.subarray(dataStart, dataStart + compressedSize);

          let xmlString = '';
          if (compressionMethod === 0) {
            xmlString = decoder.decode(compressedData);
          } else if (compressionMethod === 8 && typeof DecompressionStream !== 'undefined') {
            const ds = new DecompressionStream('deflate-raw');
            const writer = ds.writable.getWriter();
            writer.write(compressedData);
            writer.close();
            const response = new Response(ds.readable);
            const decompressedBuffer = await response.arrayBuffer();
            xmlString = decoder.decode(decompressedBuffer);
          } else {
            throw new Error('Trình duyệt không hỗ trợ giải nén luồng dữ liệu này.');
          }

          return DocumentExtractor.convertWordXmlToPlainText(xmlString);
        }

        ptr += 46 + fileNameLen + extraLen + commentLen;
      }

      throw new Error('Không tìm thấy thành phần word/document.xml bên trong file .docx.');
    }

    static convertWordXmlToPlainText(xmlString) {
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlString, 'application/xml');
      const paragraphs = xmlDoc.getElementsByTagNameNS('*', 'p');
      const lines = [];

      for (let i = 0; i < paragraphs.length; i++) {
        const p = paragraphs[i];
        const texts = p.getElementsByTagNameNS('*', 't');
        let lineText = '';
        for (let j = 0; j < texts.length; j++) {
          lineText += texts[j].textContent || '';
        }
        if (lineText.trim()) {
          lines.push(lineText.trim());
        }
      }
      return lines.join('\n');
    }
  }

  /* ==========================================================================
     3. QUESTION PARSER & STRICT VALIDATOR (XỬ LÝ ĐÚNG CHUẨN FILE MẪU)
     ========================================================================== */
  class QuizParser {
    /**
     * Phân tích văn bản thô thành danh sách câu hỏi chuẩn hóa và danh sách lỗi
     */
    static parse(rawText, ext = 'docx') {
      if (ext === 'json') {
        return QuizParser.parseJsonFormat(rawText);
      }

      // Tiền xử lý văn bản:
      // 1. Chuẩn hóa dấu xuống dòng
      // 2. Tách trường hợp tiêu đề dính liền với "Câu X" trên cùng 1 dòng (VD: "Bộ câu hỏi chương 1:Câu 1. Ai là...")
      const normalizedText = rawText
        .replace(/\r\n?/g, '\n')
        .replace(/([^\n])(?=(?:Câu|Question)\s+\d+\s*[.:])/gi, '$1\n');

      const rawLines = normalizedText
        .split('\n')
        .map((line) => line.replace(/\u00A0/g, ' ').trim())
        .filter((line) => line.length > 0);

      let currentSection = 'Phần chung';
      const sectionsSet = new Set();
      const rawQuestionBlocks = [];
      let currentBlock = null;

      // Biểu thức chính quy nhận diện các thành phần:
      const sectionRegex = /^['"‘’“”`]*\s*(Phần\s+\d+\s*[:.-]?\s*.+|Chương\s+\d+\s*[:.-]?\s*.+|Bộ câu hỏi\s+chương\s+\d+\s*[:.-]?.*)$/i;
      const questionStartRegex = /^(Câu\s+\d+|Question\s+\d+)\s*[.:]\s*(.+)$/i;
      const optionLineRegex = /^(\*)?\s*([A-D])\s*[.:)]\s*(.+)$/i;
      const explanationLineRegex = /^(?:Giải thích|Explanation|Hướng dẫn)\s*[:.-]\s*(.+)$/i;

      for (const line of rawLines) {
        // 1. Kiểm tra dòng câu hỏi mới (Ưu tiên kiểm tra trước để tránh nhầm với tiêu đề)
        const qMatch = line.match(questionStartRegex);
        if (qMatch) {
          if (currentBlock) {
            rawQuestionBlocks.push(currentBlock);
          }
          const originalLabel = qMatch[1].replace(/\s+/g, ' ').trim();
          const questionBody = qMatch[2].trim();

          currentBlock = {
            originalLabel,
            section: currentSection,
            textLines: [questionBody],
            options: [],
            explanationLines: [],
            lastTarget: 'question'
          };
          sectionsSet.add(currentSection);
          continue;
        }

        // 2. Kiểm tra dòng tiêu đề Phần / Chương
        const secMatch = line.match(sectionRegex);
        if (secMatch) {
          const cleanedSec = secMatch[1].replace(/[:.]$/, '').trim();
          if (cleanedSec) {
            currentSection = cleanedSec;
          }
          continue;
        }

        // 3. Nếu đang trong một câu hỏi, kiểm tra dòng Đáp án A, B, C, D (có hoặc không có dấu *)
        if (currentBlock) {
          const optMatch = line.match(optionLineRegex);
          if (optMatch) {
            const isCorrect = Boolean(optMatch[1] === '*');
            const letter = optMatch[2].toUpperCase();
            let optText = optMatch[3].trim();

            // Trường hợp dấu * đặt ở đầu nội dung sau chữ A. (VD: "A. *Nội dung")
            let starredInside = false;
            if (!isCorrect && optText.startsWith('*')) {
              optText = optText.replace(/^\*+\s*/, '');
              starredInside = true;
            }

            currentBlock.options.push({
              originalLetter: letter,
              text: optText,
              isCorrect: isCorrect || starredInside
            });
            currentBlock.lastTarget = 'option';
            continue;
          }

          // 4. Kiểm tra dòng Giải thích (nếu có trong tài liệu gốc)
          const expMatch = line.match(explanationLineRegex);
          if (expMatch) {
            currentBlock.explanationLines.push(expMatch[1].trim());
            currentBlock.lastTarget = 'explanation';
            continue;
          }

          // 5. Dòng văn bản nối tiếp (câu hỏi nhiều dòng hoặc đáp án nhiều dòng)
          if (currentBlock.lastTarget === 'question') {
            currentBlock.textLines.push(line);
          } else if (currentBlock.lastTarget === 'explanation') {
            currentBlock.explanationLines.push(line);
          } else if (
            currentBlock.lastTarget === 'option' &&
            currentBlock.options.length > 0
          ) {
            const lastOpt = currentBlock.options[currentBlock.options.length - 1];
            lastOpt.text += ' ' + line;
          }
        }
      }

      if (currentBlock) {
        rawQuestionBlocks.push(currentBlock);
      }

      // Kiểm duyệt nghiêm ngặt từng câu hỏi (Strict Validation)
      return QuizParser.validateBlocks(rawQuestionBlocks);
    }

    static validateBlocks(blocks) {
      const validQuestions = [];
      const invalidWarnings = [];
      const validSections = new Set();

      blocks.forEach((block, idx) => {
        const questionTitle = block.textLines.join(' ').trim();
        const refName = `${block.originalLabel} (${block.section})`;

        if (!questionTitle) {
          invalidWarnings.push(`${refName}: Nội dung câu hỏi bị trống.`);
          return;
        }

        if (block.options.length !== 4) {
          invalidWarnings.push(
            `${refName}: Có ${block.options.length} lựa chọn (yêu cầu đúng 4 lựa chọn A, B, C, D).`
          );
          return;
        }

        const emptyOpt = block.options.find((o) => !o.text || !o.text.trim());
        if (emptyOpt) {
          invalidWarnings.push(`${refName}: Có lựa chọn đáp án bị bỏ trống nội dung.`);
          return;
        }

        const correctOptions = block.options.filter((o) => o.isCorrect);
        if (correctOptions.length === 0) {
          // Tuyệt đối không tự bịa đáp án đúng khi file nguồn thiếu dấu *
          invalidWarnings.push(
            `${refName}: Chưa đánh dấu (*) cho đáp án đúng.`
          );
          return;
        }

        if (correctOptions.length > 1) {
          invalidWarnings.push(
            `${refName}: Có tới ${correctOptions.length} đáp án được đánh dấu (*), chỉ cho phép 1 đáp án đúng.`
          );
          return;
        }

        // Tạo câu hỏi hợp lệ với ID bất biến
        const globalIndex = validQuestions.length + 1;
        const qId = `q_${String(globalIndex).padStart(4, '0')}`;

        const normalizedOptions = block.options.map((opt, optIdx) => ({
          id: `${qId}_opt_${optIdx}`,
          originalLetter: opt.originalLetter,
          text: opt.text.trim(),
          isCorrect: opt.isCorrect
        }));

        validQuestions.push({
          id: qId,
          globalIndex,
          originalLabel: block.originalLabel,
          section: block.section,
          text: questionTitle,
          explanation:
            block.explanationLines.length > 0
              ? block.explanationLines.join(' ').trim()
              : null,
          options: normalizedOptions
        });

        validSections.add(block.section);
      });

      return {
        validQuestions,
        invalidWarnings,
        sections: Array.from(validSections)
      };
    }

    static parseJsonFormat(rawText) {
      let data;
      try {
        data = JSON.parse(rawText);
      } catch (_) {
        throw new Error('File JSON không đúng cú pháp.');
      }

      const list = Array.isArray(data) ? data : data.questions;
      if (!Array.isArray(list)) {
        throw new Error('Cấu trúc JSON phải chứa một mảng câu hỏi.');
      }

      const blocks = list.map((item, i) => ({
        originalLabel: item.originalLabel || `Câu ${i + 1}`,
        section: item.section || 'Phần chung',
        textLines: [item.text || item.question || ''],
        explanationLines: item.explanation ? [item.explanation] : [],
        options: Array.isArray(item.options)
          ? item.options.map((o, idx) => ({
            originalLetter: ['A', 'B', 'C', 'D'][idx] || 'A',
            text: typeof o === 'string' ? o : o.text,
            isCorrect: Boolean(o.isCorrect)
          }))
          : []
      }));

      return QuizParser.validateBlocks(blocks);
    }
  }

  /* ==========================================================================
     4. APPLICATION STATE & QUIZ ENGINE
     ========================================================================== */
  const AppState = {
    // Dữ liệu gốc từ File đã tải lên
    dataset: {
      fileName: '',
      fileSize: 0,
      fileExt: '',
      questions: [],       // Tất cả câu hỏi hợp lệ
      sections: [],        // Danh sách các Phần/Chương
      warnings: [],        // Danh sách cảnh báo các câu lỗi định dạng
      persistentFlaggedIds: new Set()
    },

    // Cấu hình do người dùng chọn
    config: {
      mode: 'instant',     // 'instant' (Quizizz Mode) | 'exam' (Exam Mode)
      section: 'all',
      rangeStart: 1,
      rangeEnd: 1,
      shuffleQuestions: false,
      shuffleOptions: false,
      autoAdvance: false
    },

    // Phiên làm bài hiện tại (Active Quiz Session)
    session: {
      active: false,
      isRetrySubset: false,
      items: [],           // Mảng các câu hỏi trong phiên kèm displayOptions đã cố định thứ tự
      currentIndex: 0,
      userAnswers: new Map(), // Map<questionId, selectedOptionId>
      flaggedIds: new Set(),  // Set<questionId>
      streak: 0,
      maxStreak: 0,
      liveCorrectCount: 0,
      elapsedSeconds: 0,
      timerInterval: null,
      autoAdvanceTimeout: null,
      navFilter: 'all',    // 'all' | 'unanswered' | 'flagged'
      reviewFilter: 'all', // 'all' | 'wrong' | 'skipped' | 'correct' | 'flagged'
      reviewSearchQuery: ''
    }
  };

  /* ==========================================================================
     5. UTILITY FUNCTIONS (SHUFFLE, FORMATTERS, DOM HELPERS)
     ========================================================================== */
  /**
   * Thuật toán xáo trộn Fisher-Yates (tạo bản sao mới, không làm biến đổi mảng gốc)
   */
  // Lấy khóa định danh độc nhất cho từng file
  function getFileStorageKey() {
    return `nebula_flags_${AppState.dataset.fileName}_${AppState.dataset.fileSize}`;
  }

  // Lưu danh sách câu được gắn nhãn vào LocalStorage
  function savePersistentFlags() {
    if (!AppState.dataset.fileName) return;
    const idsArray = Array.from(AppState.dataset.persistentFlaggedIds);
    localStorage.setItem(getFileStorageKey(), JSON.stringify(idsArray));
  }

  // Nạp danh sách câu được gắn nhãn từ LocalStorage khi nạp file
  function loadPersistentFlags() {
    if (!AppState.dataset.fileName) return;

    AppState.dataset.persistentFlaggedIds = new Set();
    const raw = localStorage.getItem(getFileStorageKey());
    if (raw) {
      try {
        const ids = JSON.parse(raw);
        if (!Array.isArray(ids)) throw new Error('Dữ liệu nhãn câu hỏi không hợp lệ.');
        AppState.dataset.persistentFlaggedIds = new Set(ids);
      } catch (e) {
        console.error('Lỗi khi nạp nhãn câu hỏi', e);
      }
      return;
    }

    // Chuyển các nhãn đã lưu bởi phiên bản cũ sang kho nhãn riêng.
    const legacyKey = `nebula_state_${AppState.dataset.fileName}_${AppState.dataset.fileSize}`;
    const legacyRaw = localStorage.getItem(legacyKey);
    if (!legacyRaw) return;

    try {
      const parsed = JSON.parse(legacyRaw);
      if (Array.isArray(parsed.flaggedIds)) {
        AppState.dataset.persistentFlaggedIds = new Set(parsed.flaggedIds);
        savePersistentFlags();
      }
    } catch (e) {
      console.error('Lỗi khi nạp nhãn câu hỏi cũ', e);
    }
  }

  // --- TÍNH NĂNG AUTO-SAVE TIẾN ĐỘ ---
  function saveSessionState() {
    if (!AppState.dataset.fileName || !AppState.session.active) return;
    const fileKey = `nebula_state_${AppState.dataset.fileName}_${AppState.dataset.fileSize}`;
    const stateToSave = {
      userAnswers: Array.from(AppState.session.userAnswers.entries())
    };
    localStorage.setItem(fileKey, JSON.stringify(stateToSave));
  }

  function restoreSessionState() {
    if (!AppState.dataset.fileName) return;
    const fileKey = `nebula_state_${AppState.dataset.fileName}_${AppState.dataset.fileSize}`;
    const savedData = localStorage.getItem(fileKey);
    if (savedData) {
      try {
        const parsed = JSON.parse(savedData);
        AppState.session.userAnswers = new Map(parsed.userAnswers || []);
        showToast('Đã khôi phục tiến độ cũ!', 'info');
      } catch (e) {
        console.error('Lỗi đọc bộ nhớ', e);
      }
    }
  }

  function shuffleArray(array) {
    const copy = [...array];
    for (let i = copy.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [copy[i], copy[j]] = [copy[j], copy[i]];
    }
    return copy;
  }

  function formatTime(totalSeconds) {
    const mins = Math.floor(totalSeconds / 60);
    const secs = totalSeconds % 60;
    return `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  }

  function formatFileSize(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    const kb = bytes / 1024;
    if (kb < 1024) return `${kb.toFixed(1)} KB`;
    return `${(kb / 1024).toFixed(2)} MB`;
  }

  function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(8px)';
      toast.style.transition = 'all 0.25s ease';
      setTimeout(() => toast.remove(), 260);
    }, 3400);
  }

  /* ==========================================================================
     6. DOM REFERENCES & INITIALIZATION
     ========================================================================== */
  const DOM = {};

  function cacheDomElements() {
    const ids = [
      // Header & Global
      'brandHomeBtn', 'headerDatasetBadge', 'headerFileName', 'headerQuestionCount',
      'openShortcutsBtn', 'soundToggleBtn', 'themeToggleBtn',
      // Views
      'viewSetup', 'viewQuiz', 'viewResults',
      // Upload Module
      'fileInput', 'uploadDropzone', 'addFileBtn', 'openFormatGuideBtn',
      'uploadLoadingState', 'loadingStepText',
      'uploadErrorState', 'uploadErrorTitle', 'uploadErrorMessage', 'retryUploadBtn',
      'selectedFilePanel', 'fileExtBadge', 'selectedFileName', 'selectedFileMeta',
      'removeFileBtn', 'parsedValidCount', 'parsedSectionCount', 'parsedInvalidCount',
      'validationWarningsBox', 'toggleWarningListBtn', 'validationWarningList',
      // Config Module
      'quizConfigCard', 'quizConfigForm', 'sectionSelect',
      'rangeAvailableHint', 'rangeStartInput', 'rangeEndInput', 'resetRangeBtn', 'rangeErrorText',
      'shuffleQuestionsToggle', 'shuffleOptionsToggle', 'autoAdvanceToggle',
      'readySummaryText', 'startQuizBtn',
      // Quiz Arena
      'quitQuizBtn', 'hudModeBadge', 'hudProgressCounter', 'hudAnsweredCounter',
      'hudProgressBarTrack', 'hudProgressFill',
      'hudStreakPill', 'hudStreakValue', 'hudScorePill', 'hudLiveScore', 'hudTimerValue',
      'questionCard', 'qCurrentIndexBadge', 'qOriginalNumberBadge', 'qSectionBadge',
      'flagQuestionBtn', 'questionText', 'optionsGrid',
      'questionFeedbackBanner', 'feedbackStatusBadge', 'feedbackCorrectText',
      'sourceExplanationBox', 'sourceExplanationText',
      'prevQuestionBtn', 'nextQuestionBtn',
      'unansweredCountBadge', 'questionPaletteGrid', 'submitQuizBtn',
      // Results & Review
      'resultScoreRing', 'resultPercentValue', 'resultGradeTag', 'resultSessionMeta',
      'resultsFeedbackText', 'metricCorrectCount', 'metricWrongCount', 'metricSkippedCount',
      'metricTimeAndStreak', 'practiceIncorrectBtn', 'practiceIncorrectBtnText',
      'retakeFullQuizBtn', 'backToSetupBtn',
      'reviewSearchInput', 'reviewListContainer',
      // Modals
      'confirmSubmitModal', 'confirmSubmitMessage', 'cancelSubmitBtn', 'confirmSubmitFinalBtn',
      'formatGuideModal', 'shortcutsModal'
    ];

    ids.forEach((id) => {
      DOM[id] = document.getElementById(id);
    });
  }

  /* ==========================================================================
     7. VIEW NAVIGATION & UPLOAD FLOW HANDLERS
     ========================================================================== */
  function switchView(targetViewName) {
    const views = {
      setup: DOM.viewSetup,
      quiz: DOM.viewQuiz,
      results: DOM.viewResults
    };

    Object.entries(views).forEach(([name, el]) => {
      if (!el) return;
      if (name === targetViewName) {
        el.hidden = false;
        el.classList.add('active');
      } else {
        el.hidden = true;
        el.classList.remove('active');
      }
    });

    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  function resetUploadStates() {
    DOM.uploadLoadingState.hidden = true;
    DOM.uploadErrorState.hidden = true;
  }

  async function handleFileSelection(file) {
    if (!file) return;

    resetUploadStates();
    DOM.uploadDropzone.hidden = true;
    DOM.selectedFilePanel.hidden = true;
    DOM.uploadLoadingState.hidden = false;

    try {
      const { ext, rawText } = await DocumentExtractor.extractContent(file, (stepMsg) => {
        DOM.loadingStepText.textContent = stepMsg;
      });

      DOM.loadingStepText.textContent = 'Đang phân tích câu hỏi và kiểm tra tính hợp lệ...';
      const parsed = QuizParser.parse(rawText, ext);

      if (parsed.validQuestions.length === 0) {
        const firstReason =
          parsed.invalidWarnings.length > 0
            ? ` Chi tiết: ${parsed.invalidWarnings[0]}`
            : ' Hãy đảm bảo mỗi câu bắt đầu bằng "Câu X." và có 1 đáp án đúng đánh dấu "*".';
        throw new Error(
          `Không tìm thấy câu hỏi hợp lệ nào trong file "${file.name}".${firstReason}`
        );
      }

      // Lưu vào AppState
      AppState.dataset.fileName = file.name;
      AppState.dataset.fileSize = file.size;
      AppState.dataset.fileExt = ext.toUpperCase();
      AppState.dataset.questions = parsed.validQuestions;
      AppState.dataset.sections = parsed.sections;
      AppState.dataset.warnings = parsed.invalidWarnings;

      // Nạp lại tất cả các nhãn đã từng đánh dấu của file này:
      loadPersistentFlags();

      renderLoadedDatasetUI();
      showToast(
        `Đã nạp thành công ${parsed.validQuestions.length} câu hỏi từ "${file.name}"!`,
        'success'
      );
    } catch (err) {
      DOM.uploadLoadingState.hidden = true;
      DOM.uploadDropzone.hidden = false;
      DOM.uploadErrorState.hidden = false;
      DOM.uploadErrorTitle.textContent = `Không thể nạp file "${file.name}"`;
      DOM.uploadErrorMessage.textContent =
        err.message || 'Đã xảy ra lỗi trong quá trình phân tích tài liệu.';
      showToast('Lỗi định dạng hoặc không đọc được file.', 'error');
    } finally {
      // Reset giá trị input file để có thể chọn lại cùng 1 file nếu muốn
      DOM.fileInput.value = '';
    }
  }

  function renderLoadedDatasetUI() {
    const { fileName, fileSize, fileExt, questions, sections, warnings } = AppState.dataset;

    DOM.uploadLoadingState.hidden = true;
    DOM.uploadErrorState.hidden = true;
    DOM.uploadDropzone.hidden = true;
    DOM.selectedFilePanel.hidden = false;

    // Cập nhật thông tin file
    DOM.fileExtBadge.textContent = fileExt;
    DOM.selectedFileName.textContent = fileName;
    DOM.selectedFileMeta.textContent = `${formatFileSize(fileSize)} • Đã kiểm duyệt hoàn tất`;

    // Cập nhật Header Badge
    DOM.headerDatasetBadge.hidden = false;
    DOM.headerFileName.textContent = fileName;
    DOM.headerQuestionCount.textContent = `${questions.length} câu`;

    // Cập nhật 3 ô thống kê
    DOM.parsedValidCount.textContent = questions.length;
    DOM.parsedSectionCount.textContent = sections.length;
    DOM.parsedInvalidCount.textContent = warnings.length;

    // Hiển thị cảnh báo nếu có câu lỗi
    if (warnings.length > 0) {
      DOM.validationWarningsBox.hidden = false;
      DOM.validationWarningList.innerHTML = '';
      warnings.forEach((w) => {
        const li = document.createElement('li');
        li.textContent = w;
        DOM.validationWarningList.appendChild(li);
      });
    } else {
      DOM.validationWarningsBox.hidden = true;
    }

    // Mở khóa bảng Cấu hình (Bước 2)
    DOM.quizConfigCard.setAttribute('aria-disabled', 'false');
    populateSectionFilter();
    syncRangeControlsWithSection();
    DOM.startQuizBtn.disabled = false;
  }

  function clearCurrentDataset() {
    AppState.dataset = {
      fileName: '',
      fileSize: 0,
      fileExt: '',
      questions: [],
      sections: [],
      warnings: [],
      persistentFlaggedIds: new Set()
    };

    resetUploadStates();
    DOM.selectedFilePanel.hidden = true;
    DOM.uploadDropzone.hidden = false;
    DOM.headerDatasetBadge.hidden = true;

    DOM.quizConfigCard.setAttribute('aria-disabled', 'true');
    DOM.sectionSelect.innerHTML = '<option value="all">Tất cả các phần (Toàn bộ tài liệu)</option>';
    DOM.rangeStartInput.value = 1;
    DOM.rangeEndInput.value = 1;
    DOM.rangeStartInput.disabled = true;
    DOM.rangeEndInput.disabled = true;
    DOM.resetRangeBtn.disabled = true;
    DOM.rangeAvailableHint.textContent = 'Khả dụng: 0 - 0';
    DOM.readySummaryText.textContent = 'Vui lòng tải lên file câu hỏi (.docx) ở Bước 1 để bắt đầu.';
    DOM.startQuizBtn.disabled = true;

    showToast('Đã xóa file câu hỏi hiện tại.', 'info');
  }

  /* ==========================================================================
     8. CONFIGURATION & QUESTION RANGE VALIDATION (A -> B)
     ========================================================================== */
  function populateSectionFilter() {
    const { questions, sections } = AppState.dataset;
    DOM.sectionSelect.innerHTML = '';

    const allOption = document.createElement('option');
    allOption.value = 'all';
    allOption.textContent = `Tất cả các phần (Tổng cộng ${questions.length} câu)`;
    DOM.sectionSelect.appendChild(allOption);

    sections.forEach((secName) => {
      const sectionQuestions = questions.filter((q) => q.section === secName);
      if (sectionQuestions.length > 0) {
        const firstIdx = sectionQuestions[0].globalIndex;
        const lastIdx = sectionQuestions[sectionQuestions.length - 1].globalIndex;
        const opt = document.createElement('option');
        opt.value = secName;
        opt.textContent = `${secName} (${sectionQuestions.length} câu: #${firstIdx} → #${lastIdx})`;
        DOM.sectionSelect.appendChild(opt);
      }
    });
  }

  function syncRangeControlsWithSection() {
    const { questions } = AppState.dataset;
    if (questions.length === 0) return;

    const selectedSec = DOM.sectionSelect.value;
    let pool = questions;
    if (selectedSec !== 'all') {
      pool = questions.filter((q) => q.section === selectedSec);
    }

    const minIndex = pool[0].globalIndex;
    const maxIndex = pool[pool.length - 1].globalIndex;

    DOM.rangeStartInput.min = 1;
    DOM.rangeStartInput.max = questions.length;
    DOM.rangeEndInput.min = 1;
    DOM.rangeEndInput.max = questions.length;

    DOM.rangeStartInput.value = minIndex;
    DOM.rangeEndInput.value = maxIndex;

    DOM.rangeStartInput.disabled = false;
    DOM.rangeEndInput.disabled = false;
    DOM.resetRangeBtn.disabled = false;

    DOM.rangeAvailableHint.textContent = `Khả dụng: Câu #1 → #${questions.length}`;
    validateAndRefreshReadySummary();
  }

  function validateAndRefreshReadySummary() {
    const total = AppState.dataset.questions.length;
    if (total === 0) return false;

    const startVal = parseInt(DOM.rangeStartInput.value, 10);
    const endVal = parseInt(DOM.rangeEndInput.value, 10);

    if (
      Number.isNaN(startVal) ||
      Number.isNaN(endVal) ||
      startVal < 1 ||
      endVal > total ||
      startVal > endVal
    ) {
      DOM.rangeErrorText.hidden = false;
      DOM.rangeErrorText.textContent = `Vui lòng nhập khoảng câu hợp lệ từ 1 đến ${total} (Câu bắt đầu ≤ Câu kết thúc).`;
      DOM.startQuizBtn.disabled = true;
      return false;
    }

    const filteredPool = getFilteredQuestionsPool(startVal, endVal, DOM.sectionSelect.value);
    if (filteredPool.length === 0) {
      DOM.rangeErrorText.hidden = false;
      DOM.rangeErrorText.textContent = 'Không có câu hỏi nào nằm trong phạm vi lọc này.';
      DOM.startQuizBtn.disabled = true;
      return false;
    }

    DOM.rangeErrorText.hidden = true;
    DOM.startQuizBtn.disabled = false;
    DOM.readySummaryText.innerHTML = `Sẵn sàng luyện tập <strong>${filteredPool.length} câu hỏi</strong> (từ #${startVal} đến #${endVal}).`;
    return true;
  }

  function getFilteredQuestionsPool(startIdx, endIdx, sectionFilter) {
    return AppState.dataset.questions.filter((q) => {
      const inRange = q.globalIndex >= startIdx && q.globalIndex <= endIdx;
      const inSection = sectionFilter === 'all' || q.section === sectionFilter;
      return inRange && inSection;
    });
  }

  /* ==========================================================================
     9. QUIZ SESSION INITIALIZATION & INDEPENDENT SHUFFLING
     ========================================================================== */
  /**
   * Khởi tạo phiên làm bài từ một danh sách câu hỏi gốc
   * Đảm bảo:
   * - Xáo trộn câu hỏi độc lập với xáo trộn đáp án
   * - Đáp án chỉ xáo trộn 1 lần duy nhất khi tạo phiên, không bao giờ xáo trộn lại khi render
   */
  function startQuizSession(sourceQuestions, isRetrySubset = false) {
    stopSessionTimers();

    // Đọc cấu hình từ UI
    const selectedModeRadio = document.querySelector('input[name="quizMode"]:checked');
    AppState.config.mode = selectedModeRadio ? selectedModeRadio.value : 'instant';
    AppState.config.section = DOM.sectionSelect.value;
    AppState.config.rangeStart = parseInt(DOM.rangeStartInput.value, 10) || 1;
    AppState.config.rangeEnd =
      parseInt(DOM.rangeEndInput.value, 10) || AppState.dataset.questions.length;
    AppState.config.shuffleQuestions = DOM.shuffleQuestionsToggle.checked;
    AppState.config.shuffleOptions = DOM.shuffleOptionsToggle.checked;
    AppState.config.autoAdvance = DOM.autoAdvanceToggle.checked;

    // 1. Xáo trộn thứ tự câu hỏi (nếu bật)
    const orderedQuestions = AppState.config.shuffleQuestions
      ? shuffleArray(sourceQuestions)
      : [...sourceQuestions];

    // 2. Xáo trộn vị trí 4 lựa chọn đáp án (nếu bật) và cố định luôn vào session item
    const letters = ['A', 'B', 'C', 'D'];
    const sessionItems = orderedQuestions.map((q, idx) => {
      const rawOptions = AppState.config.shuffleOptions
        ? shuffleArray(q.options)
        : [...q.options];

      const displayOptions = rawOptions.map((opt, optPos) => ({
        id: opt.id,
        displayLetter: letters[optPos] || 'A',
        originalLetter: opt.originalLetter,
        text: opt.text,
        isCorrect: opt.isCorrect
      }));

      return {
        sessionOrder: idx + 1,
        question: q,
        displayOptions
      };
    });

    // Khởi tạo trạng thái phiên
    AppState.session = {
      active: true,
      isRetrySubset,
      items: sessionItems,
      currentIndex: 0,
      userAnswers: new Map(),
      // Lấy trực tiếp từ kho lưu trữ nhãn vĩnh viễn:
      flaggedIds: AppState.dataset.persistentFlaggedIds,
      streak: 0,
      maxStreak: 0,
      liveCorrectCount: 0,
      elapsedSeconds: 0,
      timerInterval: null,
      autoAdvanceTimeout: null,
      navFilter: 'all',
      reviewFilter: 'all',
      reviewSearchQuery: ''
    };

    // Cập nhật HUD Mode
    DOM.hudModeBadge.textContent =
      AppState.config.mode === 'instant' ? '⚡ Quizizz Mode' : '📝 Exam Mode';
    DOM.hudScorePill.hidden = AppState.config.mode !== 'instant';
    DOM.hudStreakPill.hidden = AppState.config.mode !== 'instant';

    // Reset các tab lọc của Navigator
    document.querySelectorAll('[data-nav-filter]').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.navFilter === 'all');
      btn.setAttribute('aria-selected', btn.dataset.navFilter === 'all' ? 'true' : 'false');
    });
    if (!isRetrySubset) {
      restoreSessionState();
    }
    startSessionTimer();
    switchView('quiz');
    renderActiveQuestion();
    renderNavigatorPalette();
    updateHudProgress();
  }

  function startSessionTimer() {
    DOM.hudTimerValue.textContent = '00:00';
    AppState.session.timerInterval = setInterval(() => {
      AppState.session.elapsedSeconds += 1;
      DOM.hudTimerValue.textContent = formatTime(AppState.session.elapsedSeconds);
    }, 1000);
  }

  function stopSessionTimers() {
    if (AppState.session.timerInterval) {
      clearInterval(AppState.session.timerInterval);
      AppState.session.timerInterval = null;
    }
    if (AppState.session.autoAdvanceTimeout) {
      clearTimeout(AppState.session.autoAdvanceTimeout);
      AppState.session.autoAdvanceTimeout = null;
    }
  }

  /* ==========================================================================
     10. QUIZ ARENA RENDERING, INTERACTION & NAVIGATION
     ========================================================================== */
  function renderActiveQuestion() {
    const { items, currentIndex, userAnswers, flaggedIds } = AppState.session;
    if (!items.length) return;

    const currentItem = items[currentIndex];
    const { question, displayOptions } = currentItem;
    const selectedOptionId = userAnswers.get(question.id) || null;
    const isAnswered = Boolean(selectedOptionId);
    const isInstantMode = AppState.config.mode === 'instant';

    // 1. Cập nhật các Badge thông tin câu hỏi (Luôn giữ số thứ tự câu gốc)
    DOM.qCurrentIndexBadge.textContent = `Câu ${currentIndex + 1} / ${items.length}`;
    DOM.qOriginalNumberBadge.textContent = `Gốc: ${question.originalLabel} (#${question.globalIndex})`;
    DOM.qSectionBadge.textContent = question.section;

    // 2. Cập nhật nút Đánh dấu (Flag)
    const isFlagged = flaggedIds.has(question.id);
    DOM.flagQuestionBtn.setAttribute('aria-pressed', isFlagged ? 'true' : 'false');
    DOM.flagQuestionBtn.querySelector('span').textContent = isFlagged
      ? 'Đã đánh dấu'
      : 'Đánh dấu';

    // 3. Nội dung câu hỏi
    DOM.questionText.textContent = question.text;

    // 4. Render 4 lựa chọn đáp án dựa trên displayOptions đã cố định
    DOM.optionsGrid.innerHTML = '';
    displayOptions.forEach((opt, idx) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'option-card';
      btn.dataset.optionId = opt.id;
      btn.dataset.optionIndex = idx;

      const keySpan = document.createElement('span');
      keySpan.className = 'option-key';
      keySpan.textContent = opt.displayLetter;

      const textSpan = document.createElement('span');
      textSpan.className = 'option-text';
      textSpan.textContent = opt.text;

      const iconSpan = document.createElement('span');
      iconSpan.className = 'option-state-icon';
      iconSpan.setAttribute('aria-hidden', 'true');

      // Xử lý trạng thái hiển thị nếu người dùng đã chọn đáp án cho câu này
      if (isAnswered) {
        if (isInstantMode) {
          // Chế độ Quizizz: Khóa nút và hiển thị rõ Đúng/Sai
          btn.disabled = true;
          if (opt.isCorrect) {
            btn.classList.add('is-correct');
            iconSpan.textContent = '✓';
          } else if (opt.id === selectedOptionId && !opt.isCorrect) {
            btn.classList.add('is-wrong');
            iconSpan.textContent = '✕';
          }
        } else {
          // Chế độ Thi thử (Exam Mode): Chỉ tô sáng đáp án đang chọn, cho phép đổi ý
          if (opt.id === selectedOptionId) {
            btn.classList.add('is-selected');
            iconSpan.textContent = '●';
          }
        }
      }

      btn.append(keySpan, textSpan, iconSpan);
      btn.addEventListener('click', () => handleSelectOption(opt.id));
      DOM.optionsGrid.appendChild(btn);
    });

    // 5. Hiển thị Hộp phản hồi & Giải thích từ tài liệu gốc (trong chế độ Quizizz)
    renderFeedbackBanner(currentItem, selectedOptionId);

    // 6. Trạng thái nút điều hướng Trước / Tiếp
    DOM.prevQuestionBtn.disabled = currentIndex === 0;
    const nextBtnSpan = DOM.nextQuestionBtn.querySelector('span');
    if (currentIndex === items.length - 1) {
      nextBtnSpan.innerHTML = 'Tới cuối bảng ';
    } else {
      nextBtnSpan.innerHTML = 'Câu tiếp <kbd>→</kbd>';
    }
  }

  function renderFeedbackBanner(currentItem, selectedOptionId) {
    const isInstantMode = AppState.config.mode === 'instant';
    if (!isInstantMode || !selectedOptionId) {
      DOM.questionFeedbackBanner.hidden = true;
      return;
    }

    const { question, displayOptions } = currentItem;
    const chosenOpt = displayOptions.find((o) => o.id === selectedOptionId);
    const correctOpt = displayOptions.find((o) => o.isCorrect);
    const isCorrect = Boolean(chosenOpt && chosenOpt.isCorrect);

    DOM.questionFeedbackBanner.hidden = false;
    DOM.questionFeedbackBanner.className = `q-feedback-banner ${isCorrect ? 'feedback-correct' : 'feedback-wrong'
      }`;

    if (isCorrect) {
      DOM.feedbackStatusBadge.textContent = '✅ Chính xác!';
      DOM.feedbackCorrectText.textContent = `Bạn đã chọn đúng đáp án ${correctOpt.displayLetter}.`;
    } else {
      DOM.feedbackStatusBadge.textContent = '❌ Chưa chính xác!';
      DOM.feedbackCorrectText.textContent = `Đáp án đúng là ${correctOpt.displayLetter}: "${correctOpt.text}"`;
    }

    // Chỉ hiển thị hộp giải thích nếu tài liệu gốc thực sự có giải thích
    if (question.explanation) {
      DOM.sourceExplanationBox.hidden = false;
      DOM.sourceExplanationText.textContent = question.explanation;
    } else {
      DOM.sourceExplanationBox.hidden = true;
    }
  }

  function handleSelectOption(optionId) {
    if (!AppState.session.active) return;

    const { items, currentIndex, userAnswers } = AppState.session;
    const currentItem = items[currentIndex];
    const { question, displayOptions } = currentItem;
    const isInstantMode = AppState.config.mode === 'instant';

    // Trong chế độ Quizizz (Instant), nếu câu này đã trả lời rồi thì không cho chọn lại
    if (isInstantMode && userAnswers.has(question.id)) {
      return;
    }

    // Hủy bộ đếm tự động chuyển câu cũ (nếu có)
    if (AppState.session.autoAdvanceTimeout) {
      clearTimeout(AppState.session.autoAdvanceTimeout);
      AppState.session.autoAdvanceTimeout = null;
    }

    // Lưu lựa chọn theo ID ổn định
    userAnswers.set(question.id, optionId);

    const chosenOpt = displayOptions.find((o) => o.id === optionId);
    const isCorrect = Boolean(chosenOpt && chosenOpt.isCorrect);

    if (isInstantMode) {
      if (isCorrect) {
        AppState.session.streak += 1;
        if (AppState.session.streak > AppState.session.maxStreak) {
          AppState.session.maxStreak = AppState.session.streak;
        }
        AppState.session.liveCorrectCount += 1;
        sfx.playCorrect();
        triggerStreakPulse();
      } else {
        AppState.session.streak = 0;
        sfx.playWrong();
      }
    } else {
      sfx.playSelect();
    }

    renderActiveQuestion();
    renderNavigatorPalette();
    updateHudProgress();
    saveSessionState();

    // Tự động chuyển sang câu tiếp theo nếu người dùng bật Auto-Advance
    if (AppState.config.autoAdvance && currentIndex < items.length - 1) {
      const delay = isInstantMode ? 1200 : 450;
      AppState.session.autoAdvanceTimeout = setTimeout(() => {
        navigateToQuestion(currentIndex + 1);
      }, delay);
    }
  }

  function triggerStreakPulse() {
    DOM.hudStreakPill.classList.remove('streak-pulse');
    void DOM.hudStreakPill.offsetWidth; // Force reflow
    DOM.hudStreakPill.classList.add('streak-pulse');
  }

  function navigateToQuestion(targetIndex) {
    const { items } = AppState.session;
    if (targetIndex < 0 || targetIndex >= items.length) return;

    if (AppState.session.autoAdvanceTimeout) {
      clearTimeout(AppState.session.autoAdvanceTimeout);
      AppState.session.autoAdvanceTimeout = null;
    }

    AppState.session.currentIndex = targetIndex;
    renderActiveQuestion();
    renderNavigatorPalette();
    updateHudProgress();
  }

  function toggleCurrentQuestionFlag() {
    const { items, currentIndex } = AppState.session;
    if (!items.length) return;

    const qId = items[currentIndex].question.id;
    const flags = AppState.dataset.persistentFlaggedIds;

    if (flags.has(qId)) {
      flags.delete(qId); // Bạn chủ động nhấn xóa thì mới mất
      showToast('Đã bỏ đánh dấu câu hỏi.', 'info');
    } else {
      flags.add(qId);    // Bạn chủ động gắn cờ
      showToast('Đã đánh dấu câu hỏi.', 'info');
    }

    // Đồng bộ với session hiện tại
    AppState.session.flaggedIds = flags;

    renderActiveQuestion();
    renderNavigatorPalette();
    savePersistentFlags(); // Lưu lại ngay lập tức
  }

  function updateHudProgress() {
    const { items, currentIndex, userAnswers, streak, liveCorrectCount } = AppState.session;
    const total = items.length;
    const answeredCount = userAnswers.size;
    const unansweredCount = total - answeredCount;
    const percent = total > 0 ? Math.round((answeredCount / total) * 100) : 0;

    DOM.hudProgressCounter.textContent = `Câu ${currentIndex + 1} / ${total}`;
    DOM.hudAnsweredCounter.textContent = `Đã trả lời: ${answeredCount}/${total}`;
    DOM.hudProgressFill.style.width = `${percent}%`;
    DOM.hudProgressBarTrack.setAttribute('aria-valuenow', String(percent));

    DOM.hudStreakValue.textContent = String(streak);
    DOM.hudLiveScore.textContent = `${liveCorrectCount}/${total}`;
    DOM.unansweredCountBadge.textContent = `Còn trống: ${unansweredCount}`;
  }

  function renderNavigatorPalette() {
    const { items, currentIndex, userAnswers, flaggedIds, navFilter } = AppState.session;
    const isInstantMode = AppState.config.mode === 'instant';

    DOM.questionPaletteGrid.innerHTML = '';

    items.forEach((item, idx) => {
      const qId = item.question.id;
      const selectedOptId = userAnswers.get(qId);
      const isAnswered = Boolean(selectedOptId);
      const isFlagged = flaggedIds.has(qId);

      // Áp dụng bộ lọc trên bảng điều hướng
      if (navFilter === 'unanswered' && isAnswered) return;
      if (navFilter === 'flagged' && !isFlagged) return;

      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'palette-btn';
      btn.textContent = String(idx + 1);
      btn.title = `Câu ${idx + 1} (${item.question.originalLabel})`;

      if (idx === currentIndex) {
        btn.classList.add('is-current');
      }

      if (isFlagged) {
        btn.classList.add('is-flagged');
      }

      if (isAnswered) {
        if (isInstantMode) {
          const chosenOpt = item.displayOptions.find((o) => o.id === selectedOptId);
          btn.classList.add(chosenOpt && chosenOpt.isCorrect ? 'is-correct' : 'is-wrong');
        } else {
          btn.classList.add('is-answered');
        }
      }

      btn.addEventListener('click', () => navigateToQuestion(idx));
      DOM.questionPaletteGrid.appendChild(btn);
    });
  }

  /* ==========================================================================
     11. SCORING, SUBMISSION & DETAILED REVIEW ENGINE
     ========================================================================== */
  function requestQuizSubmission() {
    const { items, userAnswers } = AppState.session;
    const unanswered = items.length - userAnswers.size;

    if (unanswered > 0) {
      DOM.confirmSubmitMessage.textContent = `Bạn vẫn còn ${unanswered} câu hỏi chưa chọn đáp án. Bạn có chắc chắn muốn nộp bài và chấm điểm ngay bây giờ không?`;
      openModal('confirmSubmitModal');
    } else {
      finalizeAndScoreQuiz();
    }
  }

  function finalizeAndScoreQuiz() {
    closeModal('confirmSubmitModal');
    stopSessionTimers();
    AppState.session.active = false;
    sfx.playComplete();

    const { items, userAnswers, elapsedSeconds } = AppState.session;
    let correctCount = 0;
    let wrongCount = 0;
    let skippedCount = 0;
    let currentExamStreak = 0;
    let computedMaxStreak = AppState.session.maxStreak;

    items.forEach((item) => {
      const qId = item.question.id;
      const selectedOptId = userAnswers.get(qId);
      if (!selectedOptId) {
        skippedCount += 1;
        currentExamStreak = 0;
        return;
      }

      const correctOpt = item.question.options.find((o) => o.isCorrect);
      if (correctOpt && selectedOptId === correctOpt.id) {
        correctCount += 1;
        currentExamStreak += 1;
        if (currentExamStreak > computedMaxStreak) {
          computedMaxStreak = currentExamStreak;
        }
      } else {
        wrongCount += 1;
        currentExamStreak = 0;
      }
    });

    const total = items.length;
    const percentage = total > 0 ? Math.round((correctCount / total) * 100) : 0;

    // Cập nhật Vòng tròn Điểm số SVG
    const circumference = 427.25; // 2 * PI * 68
    const offset = circumference - (percentage / 100) * circumference;
    DOM.resultScoreRing.style.strokeDashoffset = String(offset);

    if (percentage >= 80) {
      DOM.resultScoreRing.style.stroke = 'var(--success)';
      DOM.resultGradeTag.textContent = 'Xuất sắc 🌟';
      DOM.resultsFeedbackText.textContent =
        'Tuyệt vời! Bạn đã nắm rất vững kiến thức của bộ câu hỏi này.';
    } else if (percentage >= 50) {
      DOM.resultScoreRing.style.stroke = 'var(--primary)';
      DOM.resultGradeTag.textContent = 'Đạt yêu cầu 👍';
      DOM.resultsFeedbackText.textContent =
        'Kết quả khá tốt! Hãy ôn lại các câu sai bên dưới để đạt điểm tuyệt đối nhé.';
    } else {
      DOM.resultScoreRing.style.stroke = 'var(--danger)';
      DOM.resultGradeTag.textContent = 'Cần cố gắng 💪';
      DOM.resultsFeedbackText.textContent =
        'Đừng nản lòng! Hãy bấm nút "Luyện lại các câu Sai & Bỏ trống" để ghi nhớ nhanh hơn.';
    }

    DOM.resultPercentValue.textContent = `${percentage}%`;
    DOM.resultSessionMeta.textContent = `${AppState.config.mode === 'instant' ? 'Chế độ Quizizz' : 'Chế độ Thi thử'
      } • ${AppState.dataset.fileName}`;

    DOM.metricCorrectCount.textContent = `${correctCount} / ${total}`;
    DOM.metricWrongCount.textContent = String(wrongCount);
    DOM.metricSkippedCount.textContent = String(skippedCount);
    DOM.metricTimeAndStreak.textContent = `${formatTime(elapsedSeconds)} • 🔥 ${computedMaxStreak}`;

    // Nút luyện tập lại câu sai + bỏ trống
    const retryTargetCount = wrongCount + skippedCount;
    DOM.practiceIncorrectBtn.disabled = retryTargetCount === 0;
    DOM.practiceIncorrectBtnText.textContent = `Luyện lại các câu Sai & Bỏ trống (${retryTargetCount})`;

    // Reset bộ lọc Review
    AppState.session.reviewFilter = 'all';
    AppState.session.reviewSearchQuery = '';
    DOM.reviewSearchInput.value = '';
    document.querySelectorAll('[data-review-filter]').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.reviewFilter === 'all');
    });

    switchView('results');
    renderReviewList();
  }

  function renderReviewList() {
    const { items, userAnswers, flaggedIds, reviewFilter, reviewSearchQuery } = AppState.session;
    const query = reviewSearchQuery.trim().toLowerCase();

    DOM.reviewListContainer.innerHTML = '';

    const filteredItems = items.filter((item) => {
      const q = item.question;
      const selectedOptId = userAnswers.get(q.id);
      const correctOpt = q.options.find((o) => o.isCorrect);
      const isSkipped = !selectedOptId;
      const isCorrect = !isSkipped && correctOpt && selectedOptId === correctOpt.id;
      const isWrong = !isSkipped && !isCorrect;
      const isFlagged = flaggedIds.has(q.id);

      if (reviewFilter === 'wrong' && !isWrong) return false;
      if (reviewFilter === 'skipped' && !isSkipped) return false;
      if (reviewFilter === 'correct' && !isCorrect) return false;
      if (reviewFilter === 'flagged' && !isFlagged) return false;

      if (query) {
        const hayStack = `${q.originalLabel} ${q.section} ${q.text}`.toLowerCase();
        if (!hayStack.includes(query)) return false;
      }

      return true;
    });

    if (filteredItems.length === 0) {
      const emptyMsg = document.createElement('p');
      emptyMsg.style.textAlign = 'center';
      emptyMsg.style.padding = '2rem';
      emptyMsg.style.color = 'var(--text-secondary)';
      emptyMsg.textContent = 'Không có câu hỏi nào khớp với bộ lọc hiện tại.';
      DOM.reviewListContainer.appendChild(emptyMsg);
      return;
    }

    filteredItems.forEach((item) => {
      const { question, displayOptions, sessionOrder } = item;
      const selectedOptId = userAnswers.get(question.id);
      const isSkipped = !selectedOptId;
      const correctOpt = displayOptions.find((o) => o.isCorrect);
      const isCorrect = !isSkipped && correctOpt && selectedOptId === correctOpt.id;

      let statusClass = 'status-skipped';
      let statusLabel = '⚠️ Bỏ trống';
      if (isCorrect) {
        statusClass = 'status-correct';
        statusLabel = '✅ Chính xác';
      } else if (!isSkipped) {
        statusClass = 'status-wrong';
        statusLabel = '❌ Chọn sai';
      }

      const card = document.createElement('article');
      card.className = `review-item ${statusClass}`;

      const header = document.createElement('div');
      header.className = 'review-item-header';
      header.innerHTML = `
        <div class="q-badges">
          <span class="q-current-badge">Câu #${sessionOrder}</span>
          <span class="q-origin-badge">Gốc: ${question.originalLabel}</span>
          <span class="q-section-badge">${question.section}</span>
        </div>
        <strong>${statusLabel}</strong>
      `;

      const qTitle = document.createElement('h4');
      qTitle.className = 'review-item-question';
      qTitle.textContent = question.text;

      const optList = document.createElement('div');
      optList.className = 'review-options-list';

      displayOptions.forEach((opt) => {
        const optDiv = document.createElement('div');
        optDiv.className = 'review-opt';

        let tagText = '';
        if (opt.isCorrect) {
          optDiv.classList.add('opt-correct');
          tagText =
            opt.id === selectedOptId ? '✓ Bạn chọn (Đúng)' : '✓ Đáp án đúng';
        } else if (opt.id === selectedOptId) {
          optDiv.classList.add('opt-user-wrong');
          tagText = '✕ Bạn đã chọn';
        }

        const leftSpan = document.createElement('span');
        leftSpan.textContent = `${opt.displayLetter}. ${opt.text}`;

        const rightSpan = document.createElement('span');
        rightSpan.style.fontSize = '0.78rem';
        rightSpan.style.whiteSpace = 'nowrap';
        rightSpan.textContent = tagText;

        optDiv.append(leftSpan, rightSpan);
        optList.appendChild(optDiv);
      });

      card.append(header, qTitle, optList);

      if (question.explanation) {
        const expBox = document.createElement('div');
        expBox.className = 'explanation-box';
        expBox.innerHTML = `<div class="explanation-tag">📄 Giải thích từ tài liệu gốc:</div>`;
        const expP = document.createElement('p');
        expP.textContent = question.explanation;
        expBox.appendChild(expP);
        card.appendChild(expBox);
      }

      DOM.reviewListContainer.appendChild(card);
    });
  }

  /* ==========================================================================
     12. MODAL & KEYBOARD SHORTCUTS MANAGEMENT
     ========================================================================== */
  function openModal(modalId) {
    const modal = document.getElementById(modalId);
    if (!modal) return;
    modal.hidden = false;
    modal.setAttribute('aria-hidden', 'false');
  }

  function closeModal(modalId) {
    const modal = document.getElementById(modalId);
    if (!modal) return;
    modal.hidden = true;
    modal.setAttribute('aria-hidden', 'true');
  }

  function handleGlobalKeydown(e) {
    // Đóng modal khi nhấn Escape
    if (e.key === 'Escape') {
      ['confirmSubmitModal', 'formatGuideModal', 'shortcutsModal'].forEach(closeModal);
      return;
    }

    // Không kích hoạt phím tắt nếu đang gõ trong ô input/search
    const tag = (e.target && e.target.tagName) || '';
    if (['INPUT', 'TEXTAREA', 'SELECT'].includes(tag)) return;

    if (e.key === '?') {
      openModal('shortcutsModal');
      return;
    }

    // Phím tắt chỉ hoạt động khi đang trong màn hình làm bài (Quiz Arena)
    if (!AppState.session.active || DOM.viewQuiz.hidden) return;

    const keyUpper = e.key.toUpperCase();
    const { items, currentIndex } = AppState.session;
    const currentItem = items[currentIndex];
    if (!currentItem) return;

    // Chọn đáp án bằng phím 1-4 hoặc A-D
    const keyMap = { '1': 0, '2': 1, '3': 2, '4': 3, A: 0, B: 1, C: 2, D: 3 };
    if (Object.prototype.hasOwnProperty.call(keyMap, keyUpper)) {
      e.preventDefault();
      const targetOpt = currentItem.displayOptions[keyMap[keyUpper]];
      if (targetOpt) {
        handleSelectOption(targetOpt.id);
      }
      return;
    }

    // Chuyển câu Trước / Tiếp
    if (e.key === 'ArrowRight' || e.key === 'Enter') {
      e.preventDefault();
      if (currentIndex < items.length - 1) {
        navigateToQuestion(currentIndex + 1);
      }
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      if (currentIndex > 0) {
        navigateToQuestion(currentIndex - 1);
      }
    } else if (keyUpper === 'F') {
      e.preventDefault();
      toggleCurrentQuestionFlag();
    }
  }

  /* ==========================================================================
     13. EVENT BINDINGS & BOOTSTRAP
     ========================================================================== */
  function bindEvents() {
    // 1. Theme & Sound toggles
    DOM.themeToggleBtn.addEventListener('click', () => {
      const htmlEl = document.documentElement;
      const nextTheme = htmlEl.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      htmlEl.setAttribute('data-theme', nextTheme);
      showToast(
        nextTheme === 'dark' ? 'Đã chuyển sang Giao diện Tối' : 'Đã chuyển sang Giao diện Sáng',
        'info'
      );
    });

    DOM.soundToggleBtn.addEventListener('click', () => {
      sfx.enabled = !sfx.enabled;
      DOM.soundToggleBtn.setAttribute('data-sound', sfx.enabled ? 'on' : 'off');
      showToast(sfx.enabled ? 'Đã bật âm thanh phản hồi' : 'Đã tắt âm thanh phản hồi', 'info');
    });

    DOM.openShortcutsBtn.addEventListener('click', () => openModal('shortcutsModal'));
    DOM.openFormatGuideBtn.addEventListener('click', () => openModal('formatGuideModal'));

    document.querySelectorAll('[data-close-modal]').forEach((btn) => {
      btn.addEventListener('click', () => closeModal(btn.dataset.closeModal));
    });

    DOM.brandHomeBtn.addEventListener('click', () => {
      if (AppState.session.active) {
        stopSessionTimers();
        AppState.session.active = false;
      }
      switchView('setup');
    });

    // 2. Mandatory File Upload & Drag-and-Drop Events
    DOM.fileInput.addEventListener('change', (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) handleFileSelection(file);
    });

    ['dragenter', 'dragover'].forEach((evtName) => {
      DOM.uploadDropzone.addEventListener(evtName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        DOM.uploadDropzone.classList.add('drag-over');
      });
    });

    ['dragleave', 'drop'].forEach((evtName) => {
      DOM.uploadDropzone.addEventListener(evtName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        DOM.uploadDropzone.classList.remove('drag-over');
      });
    });

    DOM.uploadDropzone.addEventListener('drop', (e) => {
      const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) handleFileSelection(file);
    });

    DOM.retryUploadBtn.addEventListener('click', () => {
      resetUploadStates();
      DOM.uploadDropzone.hidden = false;
      DOM.fileInput.click();
    });

    DOM.removeFileBtn.addEventListener('click', clearCurrentDataset);

    DOM.toggleWarningListBtn.addEventListener('click', () => {
      const isHidden = DOM.validationWarningList.hidden;
      DOM.validationWarningList.hidden = !isHidden;
      DOM.toggleWarningListBtn.textContent = isHidden ? 'Ẩn chi tiết' : 'Xem chi tiết';
    });

    // 3. Quiz Configuration Events
    DOM.sectionSelect.addEventListener('change', syncRangeControlsWithSection);
    DOM.rangeStartInput.addEventListener('input', validateAndRefreshReadySummary);
    DOM.rangeEndInput.addEventListener('input', validateAndRefreshReadySummary);

    DOM.resetRangeBtn.addEventListener('click', () => {
      DOM.sectionSelect.value = 'all';
      syncRangeControlsWithSection();
    });

    DOM.quizConfigForm.addEventListener('submit', (e) => {
      e.preventDefault();
      if (!validateAndRefreshReadySummary()) return;

      const startIdx = parseInt(DOM.rangeStartInput.value, 10);
      const endIdx = parseInt(DOM.rangeEndInput.value, 10);
      const pool = getFilteredQuestionsPool(startIdx, endIdx, DOM.sectionSelect.value);
      startQuizSession(pool, false);
    });

    // 4. Quiz Arena Controls
    DOM.quitQuizBtn.addEventListener('click', () => {
      stopSessionTimers();
      AppState.session.active = false;
      switchView('setup');
    });

    DOM.prevQuestionBtn.addEventListener('click', () => {
      navigateToQuestion(AppState.session.currentIndex - 1);
    });

    DOM.nextQuestionBtn.addEventListener('click', () => {
      navigateToQuestion(AppState.session.currentIndex + 1);
    });

    DOM.flagQuestionBtn.addEventListener('click', toggleCurrentQuestionFlag);

    document.querySelectorAll('[data-nav-filter]').forEach((tabBtn) => {
      tabBtn.addEventListener('click', () => {
        AppState.session.navFilter = tabBtn.dataset.navFilter;
        document.querySelectorAll('[data-nav-filter]').forEach((b) => {
          const active = b === tabBtn;
          b.classList.toggle('active', active);
          b.setAttribute('aria-selected', active ? 'true' : 'false');
        });
        renderNavigatorPalette();
      });
    });

    DOM.submitQuizBtn.addEventListener('click', requestQuizSubmission);
    DOM.cancelSubmitBtn.addEventListener('click', () => closeModal('confirmSubmitModal'));
    DOM.confirmSubmitFinalBtn.addEventListener('click', finalizeAndScoreQuiz);

    // 5. Results & Retake Actions
    DOM.practiceIncorrectBtn.addEventListener('click', () => {
      const { items, userAnswers } = AppState.session;
      const incorrectOrSkippedQuestions = items
        .filter((item) => {
          const selectedId = userAnswers.get(item.question.id);
          const correctOpt = item.question.options.find((o) => o.isCorrect);
          return !selectedId || !correctOpt || selectedId !== correctOpt.id;
        })
        .map((item) => item.question);

      if (incorrectOrSkippedQuestions.length > 0) {
        startQuizSession(incorrectOrSkippedQuestions, true);
        showToast(
          `Bắt đầu ôn tập ${incorrectOrSkippedQuestions.length} câu Sai & Bỏ trống!`,
          'info'
        );
      }
    });

    DOM.retakeFullQuizBtn.addEventListener('click', () => {
      const startIdx = AppState.config.rangeStart;
      const endIdx = AppState.config.rangeEnd;
      const pool = getFilteredQuestionsPool(startIdx, endIdx, AppState.config.section);
      startQuizSession(pool, false);
    });

    DOM.backToSetupBtn.addEventListener('click', () => switchView('setup'));

    // 6. Review Filters & Search
    document.querySelectorAll('[data-review-filter]').forEach((pillBtn) => {
      pillBtn.addEventListener('click', () => {
        AppState.session.reviewFilter = pillBtn.dataset.reviewFilter;
        document.querySelectorAll('[data-review-filter]').forEach((b) => {
          b.classList.toggle('active', b === pillBtn);
        });
        renderReviewList();
      });
    });

    DOM.reviewSearchInput.addEventListener('input', (e) => {
      AppState.session.reviewSearchQuery = e.target.value || '';
      renderReviewList();
    });

    // 7. Global Keyboard Shortcuts
    window.addEventListener('keydown', handleGlobalKeydown);
  }

  document.addEventListener('DOMContentLoaded', () => {
    cacheDomElements();
    bindEvents();
  });
})();
