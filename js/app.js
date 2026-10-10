/**
 * MICROSOFT WORD ONLINE - CORE APPLICATION ENGINE
 * Full-featured Ribbon engine with formatting, selection preservation,
 * clipboard, typography, paragraph alignment, rulers, and export tools.
 */

document.addEventListener('DOMContentLoaded', () => {
  // ================= STATE & REFS =================
  const docContent = document.getElementById('document-content');
  const docTitleInput = document.getElementById('doc-title-input');
  const saveStatus = document.getElementById('save-status');
  const wordCountNum = document.getElementById('word-count-num');
  const charCountNum = document.getElementById('char-count-num');
  const toastContainer = document.getElementById('toast-container');
  const zoomSlider = document.getElementById('zoom-slider');
  const zoomPercentage = document.getElementById('zoom-percentage');
  const pagesContainer = document.getElementById('pages-container');

  let activeFormatPainter = null; // Stored computed styles
  let formatPainterSticky = false;
  let savedSelection = null;
  let lastHighlightColor = '#ffff00';
  let lastFontColor = '#d32f2f';
  let lastShadingColor = '#e0e0e0';

  // ================= OFFICE THEME COLOR MATRIX GENERATOR =================
  const baseThemeColors = [
    '#ffffff', '#000000', '#eeece1', '#1f497d', '#4f81bd', '#c0504d', '#9bbb59', '#8064a2', '#4bacc6', '#f79646'
  ];

  const themeVariations = [
    [0.9, 0.8, 0.6, 0.4, 0.2], // Tints and shades
  ];

  function generateThemeColorPalette(containerId, callback) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = '';

    // Standard Office Matrix: 10 base colors + 5 rows of shades/tints
    const shadesMatrix = [
      ['#ffffff', '#000000', '#eeece1', '#1f497d', '#4f81bd', '#c0504d', '#9bbb59', '#8064a2', '#4bacc6', '#f79646'],
      ['#f2f2f2', '#7f7f7f', '#ddd9c3', '#c6d9f1', '#dbe5f1', '#f2dcdb', '#ebf1dd', '#e5dfec', '#dbeef3', '#fdeada'],
      ['#d8d8d8', '#595959', '#c4bd97', '#8db3e2', '#b8cce4', '#e5b9b7', '#d7e3bc', '#ccc1da', '#b7dde8', '#fbd5b5'],
      ['#bfbfbf', '#3f3f3f', '#948a54', '#548dd4', '#95b3d7', '#d99694', '#c3d69b', '#b2a2c7', '#92cddc', '#fac08f'],
      ['#a5a5a5', '#262626', '#494529', '#17365d', '#366092', '#953734', '#76933c', '#604a7b', '#31859b', '#e36c09'],
      ['#7f7f7f', '#0c0c0c', '#1d1b10', '#0f243e', '#244062', '#632423', '#4f6128', '#403151', '#205867', '#974806']
    ];

    shadesMatrix.forEach(row => {
      row.forEach(hex => {
        const cell = document.createElement('div');
        cell.className = 'color-matrix-cell';
        cell.style.backgroundColor = hex;
        cell.title = hex;
        cell.addEventListener('mousedown', (e) => {
          e.preventDefault();
          callback(hex);
          closeAllPopovers();
        });
        container.appendChild(cell);
      });
    });
  }

  generateThemeColorPalette('theme-colors-grid', (color) => {
    applyFontColor(color);
  });

  generateThemeColorPalette('shading-colors-grid', (color) => {
    applyParagraphShading(color);
  });

  // ================= UTILITIES: SELECTION & COMMANDS =================
  function saveCurrentSelection() {
    const sel = window.getSelection();
    if (sel.getRangeAt && sel.rangeCount) {
      savedSelection = sel.getRangeAt(0).cloneRange();
    }
  }

  function restoreSelection() {
    if (savedSelection) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(savedSelection);
    }
  }

  function execCmd(command, value = null) {
    docContent.focus();
    document.execCommand(command, false, value);
    updateRibbonState();
    triggerAutoSave();
    updateWordMetrics();
  }

  function showToast(msg, icon = 'fa-solid fa-circle-check') {
    const toast = document.createElement('div');
    toast.className = 'office-toast';
    toast.innerHTML = `<i class="${icon} text-blue"></i> <span>${msg}</span>`;
    toastContainer.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 2800);
  }

  // ================= POPOVERS & DROPDOWNS HANDLING =================
  function closeAllPopovers() {
    document.querySelectorAll('.dropdown-popover, .select-items').forEach(el => {
      el.classList.remove('show');
      el.classList.add('select-hide');
    });
  }

  document.addEventListener('click', (e) => {
    if (!e.target.closest('.split-button-container') && 
        !e.target.closest('.dropdown-btn-container') && 
        !e.target.closest('.custom-select-container') &&
        !e.target.closest('.color-split')) {
      closeAllPopovers();
    }
  });

  function togglePopover(popoverEl) {
    const isShowing = popoverEl.classList.contains('show') || !popoverEl.classList.contains('select-hide');
    closeAllPopovers();
    if (!isShowing) {
      popoverEl.classList.add('show');
      popoverEl.classList.remove('select-hide');
    }
  }

  // ================= RIBBON TABS SWITCHER =================
  const ribbonTabs = document.querySelectorAll('.ribbon-tab');
  const ribbonPanels = document.querySelectorAll('.ribbon-content');

  ribbonTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const targetTab = tab.dataset.tab;
      
      // If File tab clicked, show fullscreen file overlay
      if (targetTab === 'file') {
        const fileRibbon = document.getElementById('ribbon-file');
        fileRibbon.style.display = 'block';
        return;
      }

      ribbonTabs.forEach(t => t.classList.remove('active'));
      ribbonPanels.forEach(p => {
        p.classList.remove('active');
        if (p.id === 'ribbon-file') p.style.display = 'none';
      });

      tab.classList.add('active');
      const targetPanel = document.getElementById(`ribbon-${targetTab}`);
      if (targetPanel) targetPanel.classList.add('active');
    });
  });

  // Back from File menu
  const btnFileBack = document.getElementById('btn-file-back');
  if (btnFileBack) {
    btnFileBack.addEventListener('click', () => {
      document.getElementById('ribbon-file').style.display = 'none';
      document.getElementById('tab-home').click();
    });
  }

  // Ribbon Collapse Button
  const toggleRibbonBtn = document.getElementById('toggle-ribbon-btn');
  const ribbonPanel = document.getElementById('ribbon-panel');
  toggleRibbonBtn.addEventListener('click', () => {
    ribbonPanel.classList.toggle('collapsed');
    toggleRibbonBtn.classList.toggle('fa-chevron-down');
    toggleRibbonBtn.classList.toggle('fa-chevron-up');
  });

  // ================= 1. CLIPBOARD & FORMAT PAINTER =================
  document.getElementById('btn-cut').addEventListener('click', () => execCmd('cut'));
  document.getElementById('btn-copy').addEventListener('click', () => execCmd('copy'));
  
  document.getElementById('btn-paste').addEventListener('click', async () => {
    try {
      const text = await navigator.clipboard.readText();
      execCmd('insertText', text);
      showToast('Đã dán nội dung từ Clipboard');
    } catch (err) {
      execCmd('paste');
    }
  });

  // Format Painter implementation
  const formatPainterBtn = document.getElementById('btn-format-painter');
  
  function captureFormat() {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    const node = sel.anchorNode.nodeType === 3 ? sel.anchorNode.parentNode : sel.anchorNode;
    const computed = window.getComputedStyle(node);
    activeFormatPainter = {
      fontFamily: computed.fontFamily,
      fontSize: computed.fontSize,
      fontWeight: computed.fontWeight,
      fontStyle: computed.fontStyle,
      textDecoration: computed.textDecoration,
      color: computed.color,
      backgroundColor: computed.backgroundColor
    };
    document.body.classList.add('format-painter-active');
    formatPainterBtn.classList.add('active');
    showToast('Đã sao chép định dạng! Hãy quét chuột qua văn bản cần dán.');
  }

  formatPainterBtn.addEventListener('click', (e) => {
    captureFormat();
  });

  docContent.addEventListener('mouseup', () => {
    if (activeFormatPainter) {
      const sel = window.getSelection();
      if (sel.rangeCount && !sel.isCollapsed) {
        const span = document.createElement('span');
        span.style.fontFamily = activeFormatPainter.fontFamily;
        span.style.fontSize = activeFormatPainter.fontSize;
        span.style.fontWeight = activeFormatPainter.fontWeight;
        span.style.fontStyle = activeFormatPainter.fontStyle;
        span.style.textDecoration = activeFormatPainter.textDecoration;
        span.style.color = activeFormatPainter.color;
        if (activeFormatPainter.backgroundColor && activeFormatPainter.backgroundColor !== 'rgba(0, 0, 0, 0)') {
          span.style.backgroundColor = activeFormatPainter.backgroundColor;
        }

        const range = sel.getRangeAt(0);
        try {
          const contents = range.extractContents();
          span.appendChild(contents);
          range.insertNode(span);
        } catch(e) {
          console.error(e);
        }

        if (!formatPainterSticky) {
          activeFormatPainter = null;
          document.body.classList.remove('format-painter-active');
          formatPainterBtn.classList.remove('active');
        }
        triggerAutoSave();
      }
    }
  });

  // ================= 2. FONT CONTROLS =================
  // Font Family Dropdown
  const fontFamilyWrapper = document.getElementById('font-family-wrapper');
  const currentFontName = document.getElementById('current-font-name');
  const fontFamilyDropdown = document.getElementById('font-family-dropdown');

  currentFontName.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(fontFamilyDropdown);
  });

  document.querySelectorAll('.font-item').forEach(item => {
    item.addEventListener('click', (e) => {
      const font = item.dataset.font;
      currentFontName.querySelector('.font-display').textContent = font;
      execCmd('fontName', font);
      closeAllPopovers();
    });
  });

  // Font Size input & dropdown
  const fontSizeInput = document.getElementById('font-size-input');
  const fontSizeArrow = document.getElementById('font-size-arrow');
  const fontSizeDropdown = document.getElementById('font-size-dropdown');

  fontSizeArrow.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(fontSizeDropdown);
  });

  function applyCustomFontSize(ptSize) {
    saveCurrentSelection();
    restoreSelection();
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    
    // Modern CSS font-size application
    const span = document.createElement('span');
    span.style.fontSize = `${ptSize}pt`;
    const range = sel.getRangeAt(0);
    const contents = range.extractContents();
    span.appendChild(contents);
    range.insertNode(span);
    fontSizeInput.value = ptSize;
    triggerAutoSave();
  }

  document.querySelectorAll('.size-item').forEach(item => {
    item.addEventListener('click', () => {
      const size = item.dataset.size;
      fontSizeInput.value = size;
      applyCustomFontSize(size);
      closeAllPopovers();
    });
  });

  fontSizeInput.addEventListener('change', () => {
    const val = parseInt(fontSizeInput.value, 10);
    if (val > 0 && val <= 144) {
      applyCustomFontSize(val);
    }
  });

  // Increase & Decrease Font Size
  document.getElementById('btn-increase-font').addEventListener('click', () => {
    let current = parseInt(fontSizeInput.value, 10) || 11;
    let next = current >= 12 ? (current >= 28 ? current + 4 : current + 2) : current + 1;
    if (next > 144) next = 144;
    applyCustomFontSize(next);
  });

  document.getElementById('btn-decrease-font').addEventListener('click', () => {
    let current = parseInt(fontSizeInput.value, 10) || 11;
    let next = current > 12 ? (current > 28 ? current - 4 : current - 2) : Math.max(1, current - 1);
    applyCustomFontSize(next);
  });

  // Change Case (Aa)
  const btnChangeCase = document.getElementById('btn-change-case');
  const caseMenu = document.getElementById('case-menu');
  btnChangeCase.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(caseMenu);
  });

  document.querySelectorAll('#case-menu .menu-item').forEach(item => {
    item.addEventListener('click', () => {
      const caseType = item.dataset.case;
      const sel = window.getSelection();
      if (!sel.rangeCount || sel.isCollapsed) return;
      const text = sel.toString();
      let transformed = text;

      switch(caseType) {
        case 'sentence':
          transformed = text.toLowerCase().replace(/(^\s*\w|[.!?]\s*\w)/g, c => c.toUpperCase());
          break;
        case 'lower':
          transformed = text.toLowerCase();
          break;
        case 'upper':
          transformed = text.toUpperCase();
          break;
        case 'capitalize':
          transformed = text.replace(/\b\w/g, c => c.toUpperCase());
          break;
        case 'toggle':
          transformed = text.split('').map(c => c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase()).join('');
          break;
      }
      execCmd('insertText', transformed);
      closeAllPopovers();
    });
  });

  // Clear Formatting
  document.getElementById('btn-clear-format').addEventListener('click', () => {
    execCmd('removeFormat');
    // Also clean up inline styling spans
    const sel = window.getSelection();
    if (sel.rangeCount && !sel.isCollapsed) {
      const range = sel.getRangeAt(0);
      const clone = range.cloneContents();
      const div = document.createElement('div');
      div.appendChild(clone);
      const cleanText = div.innerText || div.textContent;
      execCmd('insertText', cleanText);
    }
    showToast('Đã xóa toàn bộ định dạng!');
  });

  // Bold, Italic, Strikethrough, Subscript, Superscript
  document.getElementById('btn-bold').addEventListener('click', () => execCmd('bold'));
  document.getElementById('btn-italic').addEventListener('click', () => execCmd('italic'));
  document.getElementById('btn-strikethrough').addEventListener('click', () => execCmd('strikeThrough'));
  document.getElementById('btn-subscript').addEventListener('click', () => execCmd('subscript'));
  document.getElementById('btn-superscript').addEventListener('click', () => execCmd('superscript'));

  // Underline & Underline Styles
  const btnUnderline = document.getElementById('btn-underline');
  const btnUnderlineStyles = document.getElementById('btn-underline-styles');
  const underlineMenu = document.getElementById('underline-menu');

  btnUnderline.addEventListener('click', () => execCmd('underline'));
  btnUnderlineStyles.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(underlineMenu);
  });

  document.querySelectorAll('.u-style-item').forEach(item => {
    item.addEventListener('click', () => {
      const style = item.dataset.ustyle;
      applyUnderlineStyle(style);
      closeAllPopovers();
    });
  });

  function applyUnderlineStyle(styleName, color = 'currentColor') {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    const span = document.createElement('span');
    span.style.textDecoration = 'underline';
    if (styleName === 'double') span.style.textDecorationStyle = 'double';
    else if (styleName === 'thick') { span.style.textDecoration = 'underline'; span.style.textDecorationThickness = '3px'; }
    else if (styleName === 'dotted') span.style.textDecorationStyle = 'dotted';
    else if (styleName === 'dashed') span.style.textDecorationStyle = 'dashed';
    else if (styleName === 'wavy') span.style.textDecorationStyle = 'wavy';
    if (color !== 'currentColor') span.style.textDecorationColor = color;

    const range = sel.getRangeAt(0);
    const contents = range.extractContents();
    span.appendChild(contents);
    range.insertNode(span);
    triggerAutoSave();
  }

  document.querySelectorAll('.u-color-dot').forEach(dot => {
    dot.addEventListener('click', () => {
      applyUnderlineStyle('solid', dot.dataset.ucolor);
      closeAllPopovers();
    });
  });

  // Text Effects (Typography)
  const btnTextEffects = document.getElementById('btn-text-effects');
  const textEffectsMenu = document.getElementById('text-effects-menu');
  btnTextEffects.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(textEffectsMenu);
  });

  document.querySelectorAll('#text-effects-menu .preset-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const effect = btn.dataset.effect;
      applyTextEffect(effect);
      closeAllPopovers();
    });
  });

  function applyTextEffect(effect) {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    const span = document.createElement('span');
    span.className = `effect-${effect}`;
    const range = sel.getRangeAt(0);
    const contents = range.extractContents();
    span.appendChild(contents);
    range.insertNode(span);
    triggerAutoSave();
  }

  // Highlight Color
  const btnHighlight = document.getElementById('btn-highlight');
  const btnHighlightArrow = document.getElementById('btn-highlight-arrow');
  const highlightPaletteMenu = document.getElementById('highlight-palette-menu');
  const highlightColorBar = document.getElementById('highlight-color-bar');

  btnHighlight.addEventListener('click', () => {
    applyHighlightColor(lastHighlightColor);
  });

  btnHighlightArrow.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(highlightPaletteMenu);
  });

  document.querySelectorAll('.highlight-grid .color-cell').forEach(cell => {
    cell.addEventListener('click', () => {
      const color = cell.dataset.hl;
      lastHighlightColor = color;
      highlightColorBar.style.backgroundColor = color;
      applyHighlightColor(color);
      closeAllPopovers();
    });
  });

  document.getElementById('btn-highlight-none').addEventListener('click', () => {
    applyHighlightColor('transparent');
    closeAllPopovers();
  });

  function applyHighlightColor(color) {
    if (color === 'transparent') {
      execCmd('hiliteColor', 'transparent');
    } else {
      execCmd('hiliteColor', color);
    }
  }

  // Font Color
  const btnFontColor = document.getElementById('btn-font-color');
  const btnFontColorArrow = document.getElementById('btn-font-color-arrow');
  const fontColorPaletteMenu = document.getElementById('font-color-palette-menu');
  const fontColorBar = document.getElementById('font-color-bar');
  const customFontColorInput = document.getElementById('custom-font-color-input');

  btnFontColor.addEventListener('click', () => {
    applyFontColor(lastFontColor);
  });

  btnFontColorArrow.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(fontColorPaletteMenu);
  });

  document.querySelectorAll('#font-color-palette-menu .std-color, #font-color-palette-menu .auto-color-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const color = btn.dataset.color || '#000000';
      applyFontColor(color);
      closeAllPopovers();
    });
  });

  customFontColorInput.addEventListener('input', (e) => {
    applyFontColor(e.target.value);
  });

  function applyFontColor(color) {
    lastFontColor = color;
    fontColorBar.style.backgroundColor = color;
    execCmd('foreColor', color);
  }

  // ================= 3. PARAGRAPH CONTROLS =================
  // Alignments
  document.getElementById('btn-align-left').addEventListener('click', () => execCmd('justifyLeft'));
  document.getElementById('btn-align-center').addEventListener('click', () => execCmd('justifyCenter'));
  document.getElementById('btn-align-right').addEventListener('click', () => execCmd('justifyRight'));
  document.getElementById('btn-align-justify').addEventListener('click', () => execCmd('justifyFull'));

  // Bullets & Bullet Library
  const btnBulletList = document.getElementById('btn-bullet-list');
  const btnBulletArrow = document.getElementById('btn-bullet-arrow');
  const bulletsMenu = document.getElementById('bullets-menu');

  btnBulletList.addEventListener('click', () => execCmd('insertUnorderedList'));
  btnBulletArrow.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(bulletsMenu);
  });

  document.querySelectorAll('.bullet-choice').forEach(choice => {
    choice.addEventListener('click', () => {
      execCmd('insertUnorderedList');
      const sel = window.getSelection();
      const node = sel.anchorNode?.parentElement?.closest('ul');
      if (node) {
        node.style.listStyleType = choice.dataset.bullet;
      }
      closeAllPopovers();
    });
  });

  // Numbering & Numbering Library
  const btnNumberList = document.getElementById('btn-number-list');
  const btnNumberArrow = document.getElementById('btn-number-arrow');
  const numberingMenu = document.getElementById('numbering-menu');

  btnNumberList.addEventListener('click', () => execCmd('insertOrderedList'));
  btnNumberArrow.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(numberingMenu);
  });

  document.querySelectorAll('.num-choice').forEach(choice => {
    choice.addEventListener('click', () => {
      execCmd('insertOrderedList');
      const sel = window.getSelection();
      const node = sel.anchorNode?.parentElement?.closest('ol');
      if (node) {
        node.style.listStyleType = choice.dataset.numtype;
      }
      closeAllPopovers();
    });
  });

  // Multilevel list
  const btnMultilevelList = document.getElementById('btn-multilevel-list');
  const btnMultilevelArrow = document.getElementById('btn-multilevel-arrow');
  const multilevelMenu = document.getElementById('multilevel-menu');

  btnMultilevelList.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(multilevelMenu);
  });
  btnMultilevelArrow.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(multilevelMenu);
  });

  // Indent & Outdent
  document.getElementById('btn-decrease-indent').addEventListener('click', () => execCmd('outdent'));
  document.getElementById('btn-increase-indent').addEventListener('click', () => execCmd('indent'));

  // Sort Paragraphs (A-Z)
  document.getElementById('btn-sort').addEventListener('click', () => {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    const node = sel.anchorNode.parentElement.closest('ul, ol') || docContent;
    
    if (node.tagName === 'UL' || node.tagName === 'OL') {
      const items = Array.from(node.querySelectorAll('li'));
      items.sort((a, b) => a.innerText.localeCompare(b.innerText, 'vi'));
      items.forEach(li => node.appendChild(li));
      showToast('Đã sắp xếp danh sách theo thứ tự A-Z');
    } else {
      const paragraphs = Array.from(docContent.querySelectorAll('p'));
      if (paragraphs.length > 1) {
        paragraphs.sort((a, b) => a.innerText.localeCompare(b.innerText, 'vi'));
        paragraphs.forEach(p => docContent.appendChild(p));
        showToast('Đã sắp xếp các đoạn văn theo thứ tự A-Z');
      }
    }
  });

  // Show/Hide Pilcrow ¶ Formatting marks
  const btnPilcrow = document.getElementById('btn-pilcrow');
  btnPilcrow.addEventListener('click', () => {
    docContent.classList.toggle('show-pilcrows');
    btnPilcrow.classList.toggle('active');
    const isActive = docContent.classList.contains('show-pilcrows');
    showToast(isActive ? 'Đang hiện ký hiệu định dạng ¶' : 'Đã ẩn ký hiệu định dạng ¶');
  });

  // Line & Paragraph Spacing
  const btnLineSpacing = document.getElementById('btn-line-spacing');
  const lineSpacingMenu = document.getElementById('line-spacing-menu');

  btnLineSpacing.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(lineSpacingMenu);
  });

  document.querySelectorAll('.spacing-choice').forEach(choice => {
    choice.addEventListener('click', () => {
      const spacing = choice.dataset.spacing;
      const sel = window.getSelection();
      const node = sel.anchorNode?.parentElement?.closest('p, div, li, h1, h2, h3') || docContent;
      node.style.lineHeight = spacing;
      document.querySelectorAll('.spacing-choice').forEach(c => c.classList.remove('active'));
      choice.classList.add('active');
      closeAllPopovers();
      triggerAutoSave();
    });
  });

  document.getElementById('btn-add-space-before').addEventListener('click', () => {
    const sel = window.getSelection();
    const node = sel.anchorNode?.parentElement?.closest('p, div, li, h1, h2, h3');
    if (node) node.style.marginTop = '16pt';
    closeAllPopovers();
  });

  document.getElementById('btn-add-space-after').addEventListener('click', () => {
    const sel = window.getSelection();
    const node = sel.anchorNode?.parentElement?.closest('p, div, li, h1, h2, h3');
    if (node) node.style.marginBottom = '16pt';
    closeAllPopovers();
  });

  document.getElementById('btn-remove-space-after').addEventListener('click', () => {
    const sel = window.getSelection();
    const node = sel.anchorNode?.parentElement?.closest('p, div, li, h1, h2, h3');
    if (node) node.style.marginBottom = '0';
    closeAllPopovers();
  });

  // Shading (Paragraph Background)
  const btnShading = document.getElementById('btn-shading');
  const btnShadingArrow = document.getElementById('btn-shading-arrow');
  const shadingPaletteMenu = document.getElementById('shading-palette-menu');
  const shadingColorBar = document.getElementById('shading-color-bar');

  btnShading.addEventListener('click', () => {
    applyParagraphShading(lastShadingColor);
  });

  btnShadingArrow.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(shadingPaletteMenu);
  });

  function applyParagraphShading(color) {
    lastShadingColor = color;
    shadingColorBar.style.backgroundColor = color;
    const sel = window.getSelection();
    const node = sel.anchorNode?.parentElement?.closest('p, div, blockquote, h1, h2, h3') || docContent;
    if (node && node !== docContent) {
      node.style.backgroundColor = color;
      node.style.padding = '8px 12px';
      node.style.borderRadius = '4px';
    } else {
      execCmd('hiliteColor', color);
    }
  }

  document.getElementById('btn-shading-none').addEventListener('click', () => {
    const sel = window.getSelection();
    const node = sel.anchorNode?.parentElement?.closest('p, div, blockquote, h1, h2, h3');
    if (node) {
      node.style.backgroundColor = 'transparent';
      node.style.padding = '0';
    }
    closeAllPopovers();
  });

  // Borders
  const btnBorders = document.getElementById('btn-borders');
  const bordersMenu = document.getElementById('borders-menu');

  btnBorders.addEventListener('click', (e) => {
    e.stopPropagation();
    togglePopover(bordersMenu);
  });

  document.querySelectorAll('#borders-menu .menu-item').forEach(item => {
    item.addEventListener('click', () => {
      const borderType = item.dataset.border;
      const sel = window.getSelection();
      const node = sel.anchorNode?.parentElement?.closest('p, div, blockquote, h1, h2, h3');
      
      if (borderType === 'hline') {
        execCmd('insertHorizontalRule');
      } else if (node) {
        node.style.border = 'none';
        if (borderType === 'bottom') node.style.borderBottom = '1.5px solid #333';
        else if (borderType === 'top') node.style.borderTop = '1.5px solid #333';
        else if (borderType === 'left') node.style.borderLeft = '3px solid #185abd';
        else if (borderType === 'right') node.style.borderRight = '1.5px solid #333';
        else if (borderType === 'all') node.style.border = '1px solid #333';
        else if (borderType === 'outside') {
          node.style.border = '1.5px solid #185abd';
          node.style.padding = '8px 12px';
        } else if (borderType === 'box-callout') {
          node.style.borderLeft = '4px solid #185abd';
          node.style.backgroundColor = '#f0f4f9';
          node.style.padding = '10px 14px';
          node.style.borderRadius = '3px';
        }
      }
      closeAllPopovers();
      triggerAutoSave();
    });
  });

  // ================= 4. STYLES GALLERY =================
  document.querySelectorAll('.style-card').forEach(card => {
    card.addEventListener('click', () => {
      document.querySelectorAll('.style-card').forEach(c => c.classList.remove('active'));
      card.classList.add('active');
      const styleType = card.dataset.style;

      switch(styleType) {
        case 'normal':
          execCmd('formatBlock', '<p>');
          break;
        case 'heading-1':
          execCmd('formatBlock', '<h1>');
          break;
        case 'heading-2':
          execCmd('formatBlock', '<h2>');
          break;
        case 'title':
          execCmd('formatBlock', '<h1>');
          const sel = window.getSelection();
          const node = sel.anchorNode?.parentElement?.closest('h1');
          if (node) {
            node.style.fontSize = '26pt';
            node.style.textAlign = 'center';
            node.style.color = '#185abd';
          }
          break;
        case 'subtitle':
          execCmd('formatBlock', '<p>');
          const p = window.getSelection().anchorNode?.parentElement?.closest('p');
          if (p) {
            p.style.fontSize = '12pt';
            p.style.fontStyle = 'italic';
            p.style.color = '#555';
            p.style.textAlign = 'center';
          }
          break;
        case 'quote':
          execCmd('formatBlock', '<blockquote>');
          break;
      }
    });
  });

  // ================= 5. EDITING: FIND & REPLACE =================
  const findReplaceDialog = document.getElementById('find-replace-dialog');
  const btnFind = document.getElementById('btn-find');
  const btnReplace = document.getElementById('btn-replace');
  const btnSelectAll = document.getElementById('btn-select-all');
  const btnCloseFind = document.getElementById('btn-close-find');
  const btnCancelDialog = document.getElementById('btn-cancel-dialog');
  const findInputText = document.getElementById('find-input-text');
  const replaceInputText = document.getElementById('replace-input-text');
  const replaceFieldRow = document.getElementById('replace-field-row');

  btnFind.addEventListener('click', () => {
    findReplaceDialog.classList.add('show');
    replaceFieldRow.style.display = 'none';
    document.getElementById('dtab-find').classList.add('active');
    document.getElementById('dtab-replace').classList.remove('active');
    findInputText.focus();
  });

  btnReplace.addEventListener('click', () => {
    findReplaceDialog.classList.add('show');
    replaceFieldRow.style.display = 'flex';
    document.getElementById('dtab-replace').classList.add('active');
    document.getElementById('dtab-find').classList.remove('active');
    replaceInputText.focus();
  });

  btnSelectAll.addEventListener('click', () => {
    execCmd('selectAll');
  });

  [btnCloseFind, btnCancelDialog].forEach(btn => {
    btn.addEventListener('click', () => findReplaceDialog.classList.remove('show'));
  });

  document.getElementById('btn-exec-find-next').addEventListener('click', () => {
    const q = findInputText.value;
    if (!q) return;
    const matchCase = document.getElementById('match-case-chk').checked;
    if (window.find) {
      window.find(q, matchCase, false, true);
    }
  });

  document.getElementById('btn-exec-replace').addEventListener('click', () => {
    const findStr = findInputText.value;
    const replaceStr = replaceInputText.value;
    if (!findStr) return;
    const text = docContent.innerHTML;
    docContent.innerHTML = text.replace(new RegExp(findStr, 'i'), replaceStr);
    showToast(`Đã thay thế "${findStr}" thành "${replaceStr}"`);
    triggerAutoSave();
  });

  document.getElementById('btn-exec-replace-all').addEventListener('click', () => {
    const findStr = findInputText.value;
    const replaceStr = replaceInputText.value;
    if (!findStr) return;
    const text = docContent.innerHTML;
    const regex = new RegExp(findStr, 'gi');
    const matches = (text.match(regex) || []).length;
    docContent.innerHTML = text.replace(regex, replaceStr);
    showToast(`Đã thay thế tất cả ${matches} vị trí!`);
    triggerAutoSave();
  });

  // ================= 6. INSERT TAB (TABLE, IMAGE, LINK, PAGE BREAK) =================
  // Table modal
  const tableDialog = document.getElementById('table-dialog');
  document.getElementById('btn-insert-table').addEventListener('click', () => {
    tableDialog.classList.add('show');
  });
  document.getElementById('btn-close-tbl-dialog').addEventListener('click', () => tableDialog.classList.remove('show'));
  document.getElementById('btn-cancel-tbl-dialog').addEventListener('click', () => tableDialog.classList.remove('show'));

  document.getElementById('btn-confirm-insert-tbl').addEventListener('click', () => {
    const rows = parseInt(document.getElementById('table-rows').value, 10) || 3;
    const cols = parseInt(document.getElementById('table-cols').value, 10) || 3;
    
    let html = '<table style="width:100%; border-collapse:collapse; margin: 12px 0;">';
    for (let r = 0; r < rows; r++) {
      html += '<tr>';
      for (let c = 0; c < cols; c++) {
        html += `<td style="border: 1px solid #999; padding: 8px 12px;">Ô ${r+1},${c+1}</td>`;
      }
      html += '</tr>';
    }
    html += '</table><p></p>';
    execCmd('insertHTML', html);
    tableDialog.classList.remove('show');
    showToast(`Đã tạo bảng ${rows}x${cols}`);
  });

  // Image modal
  const imageDialog = document.getElementById('image-dialog');
  document.getElementById('btn-insert-image').addEventListener('click', () => {
    imageDialog.classList.add('show');
  });
  document.getElementById('btn-close-img-dialog').addEventListener('click', () => imageDialog.classList.remove('show'));
  document.getElementById('btn-cancel-img-dialog').addEventListener('click', () => imageDialog.classList.remove('show'));

  document.getElementById('btn-confirm-insert-img').addEventListener('click', () => {
    const fileInput = document.getElementById('image-file-input');
    const urlInput = document.getElementById('image-url-input');

    if (fileInput.files && fileInput.files[0]) {
      const reader = new FileReader();
      reader.onload = (e) => {
        execCmd('insertImage', e.target.result);
        imageDialog.classList.remove('show');
        showToast('Đã chèn hình ảnh thành công');
      };
      reader.readAsDataURL(fileInput.files[0]);
    } else if (urlInput.value.trim()) {
      execCmd('insertImage', urlInput.value.trim());
      imageDialog.classList.remove('show');
      showToast('Đã chèn hình ảnh từ URL');
    }
  });

  // Link
  document.getElementById('btn-insert-link').addEventListener('click', () => {
    const url = prompt('Nhập địa chỉ liên kết (URL):', 'https://');
    if (url) execCmd('createLink', url);
  });

  // Page break
  document.getElementById('btn-insert-page-break').addEventListener('click', () => {
    const pageBreakHtml = '<div style="page-break-after: always; border-top: 2px dashed #185abd; margin: 28px 0; text-align: center; color: #185abd; font-size: 11px;">--- [Ngắt trang Page Break] ---</div><p></p>';
    execCmd('insertHTML', pageBreakHtml);
    showToast('Đã thêm ngắt trang');
  });

  // ================= 7. LAYOUT & VIEW CONTROLS =================
  // Dark mode toggle
  document.getElementById('btn-toggle-dark').addEventListener('click', () => {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    showToast(isDark ? 'Chế độ Tối (Dark Mode) đã bật' : 'Chế độ Sáng đã bật');
  });

  // Fullscreen
  document.getElementById('btn-view-fullscreen').addEventListener('click', () => {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen();
    } else {
      document.exitFullscreen();
    }
  });

  // Ruler Toggle
  document.getElementById('btn-view-ruler').addEventListener('click', function() {
    const hRuler = document.getElementById('ruler-h-container');
    const isHidden = hRuler.style.display === 'none';
    hRuler.style.display = isHidden ? 'flex' : 'none';
    this.classList.toggle('active');
    showToast(isHidden ? 'Đã hiển thị thước kẻ' : 'Đã ẩn thước kẻ');
  });

  // Quick Action: Print & Save
  document.getElementById('qa-print').addEventListener('click', () => window.print());
  document.getElementById('qa-save').addEventListener('click', () => {
    triggerAutoSave(true);
    showToast('Đã lưu tài liệu thành công!');
  });
  document.getElementById('qa-undo').addEventListener('click', () => execCmd('undo'));
  document.getElementById('qa-redo').addEventListener('click', () => execCmd('redo'));

  // Export Cards in File Menu
  document.getElementById('export-pdf').addEventListener('click', () => {
    document.getElementById('ribbon-file').style.display = 'none';
    window.print();
  });

  document.getElementById('export-docx').addEventListener('click', () => {
    const content = docContent.innerHTML;
    const title = docTitleInput.value.trim() || 'Document';
    const blob = new Blob([`
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head><meta charset='utf-8'><title>${title}</title></head>
      <body style="font-family: 'Aptos', 'Calibri', Arial, sans-serif; font-size: 11pt;">
      ${content}
      </body></html>
    `], { type: 'application/msword;charset=utf-8' });
    
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${title}.doc`;
    link.click();
    showToast(`Đang tải tệp ${title}.doc`);
  });

  document.getElementById('export-txt').addEventListener('click', () => {
    const text = docContent.innerText || docContent.textContent;
    const title = docTitleInput.value.trim() || 'Document';
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${title}.txt`;
    link.click();
    showToast(`Đã xuất văn bản ${title}.txt`);
  });

  // ================= 8. LIVE METRICS & WORD COUNT =================
  function updateWordMetrics() {
    const text = docContent.innerText || '';
    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    const chars = text.length;
    wordCountNum.textContent = words.toLocaleString();
    charCountNum.textContent = chars.toLocaleString();
  }

  docContent.addEventListener('input', () => {
    updateWordMetrics();
    triggerAutoSave();
  });

  // ================= 9. ZOOM ENGINE =================
  function applyZoom(scalePercent) {
    const scale = scalePercent / 100;
    pagesContainer.style.transform = `scale(${scale})`;
    zoomPercentage.textContent = `${scalePercent}%`;
    zoomSlider.value = scalePercent;
  }

  zoomSlider.addEventListener('input', (e) => applyZoom(parseInt(e.target.value, 10)));
  document.getElementById('btn-zoom-in').addEventListener('click', () => {
    let val = Math.min(200, parseInt(zoomSlider.value, 10) + 10);
    applyZoom(val);
  });
  document.getElementById('btn-zoom-out').addEventListener('click', () => {
    let val = Math.max(50, parseInt(zoomSlider.value, 10) - 10);
    applyZoom(val);
  });
  document.getElementById('btn-zoom-reset').addEventListener('click', () => applyZoom(100));

  // ================= 10. AUTOSAVE SYSTEM =================
  let autoSaveTimer = null;
  function triggerAutoSave(force = false) {
    if (saveStatus) saveStatus.innerHTML = '<i class="fa-solid fa-arrows-rotate fa-spin"></i> Đang lưu...';
    clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(() => {
      if (window.cloudDocs) {
        window.cloudDocs.saveActiveDocument(force);
      } else {
        localStorage.setItem('word_clone_doc', docContent.innerHTML);
        localStorage.setItem('word_clone_title', docTitleInput.value);
        if (saveStatus) saveStatus.innerHTML = '<i class="fa-solid fa-cloud-check"></i> Đã lưu vào máy';
      }
    }, force ? 50 : 600);
  }

  // ================= 11. KEYBOARD SHORTCUTS =================
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey || e.metaKey) {
      switch(e.key.toLowerCase()) {
        case 's':
          e.preventDefault();
          triggerAutoSave(true);
          showToast('Đã lưu văn bản!');
          break;
        case 'f':
          e.preventDefault();
          btnFind.click();
          break;
        case 'h':
          e.preventDefault();
          btnReplace.click();
          break;
        case 'p':
          e.preventDefault();
          window.print();
          break;
      }
    }
  });

  // Tab key indents in editor
  docContent.addEventListener('keydown', (e) => {
    if (e.key === 'Tab') {
      e.preventDefault();
      if (e.shiftKey) {
        execCmd('outdent');
      } else {
        execCmd('indent');
      }
    }
  });

  // Update Ribbon Active State based on selection
  function updateRibbonState() {
    try {
      document.getElementById('btn-bold').classList.toggle('active', document.queryCommandState('bold'));
      document.getElementById('btn-italic').classList.toggle('active', document.queryCommandState('italic'));
      document.getElementById('btn-underline').classList.toggle('active', document.queryCommandState('underline'));
      document.getElementById('btn-strikethrough').classList.toggle('active', document.queryCommandState('strikeThrough'));
      document.getElementById('btn-align-left').classList.toggle('active', document.queryCommandState('justifyLeft'));
      document.getElementById('btn-align-center').classList.toggle('active', document.queryCommandState('justifyCenter'));
      document.getElementById('btn-align-right').classList.toggle('active', document.queryCommandState('justifyRight'));
      document.getElementById('btn-align-justify').classList.toggle('active', document.queryCommandState('justifyFull'));
      document.getElementById('btn-bullet-list').classList.toggle('active', document.queryCommandState('insertUnorderedList'));
      document.getElementById('btn-number-list').classList.toggle('active', document.queryCommandState('insertOrderedList'));
    } catch(err) {}
  }

  document.addEventListener('selectionchange', () => {
    if (document.activeElement === docContent || docContent.contains(document.activeElement)) {
      updateRibbonState();
    }
  });

  // Initial calculation
  updateWordMetrics();
  window.updateWordMetrics = updateWordMetrics;

  // ================= 12. DARK MODE TOGGLE =================
  const toggleDarkMode = () => {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('clouddocs_dark_mode', isDark ? 'true' : 'false');
  };

  const savedDark = localStorage.getItem('clouddocs_dark_mode');
  if (savedDark === 'true') {
    document.body.classList.add('dark-mode');
  }

  const btnDark1 = document.getElementById('btn-toggle-dark');
  const btnDark2 = document.getElementById('btn-toggle-dark-explorer');
  if (btnDark1) btnDark1.addEventListener('click', toggleDarkMode);
  if (btnDark2) btnDark2.addEventListener('click', toggleDarkMode);
});
