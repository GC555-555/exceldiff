// =====================================================
// 样式提取模块 - 直接从xlsx文件解析样式信息
// xlsx-js-style 库不支持在浏览器端正确读取 cell.s
// 因此我们使用 JSZip 直接解析 styles.xml 和 sheet XML
// =====================================================

// 主函数：从 xlsx 原始数据中提取样式
async function extractStylesFromXlsx(fileData) {
  try {
    if (!fileData) return {};

    const zip = await JSZip.loadAsync(fileData);
    if (!zip) return {};

    // 读取 styles.xml
    const stylesXml = zip.file('xl/styles.xml');
    if (!stylesXml) {
      console.warn('未找到 styles.xml，样式可能无法提取');
      return {};
    }

    const stylesText = await stylesXml.async('text');
    if (!stylesText) return {};

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(stylesText, 'text/xml');
    if (xmlDoc.documentElement && xmlDoc.documentElement.tagName === 'parsererror') {
      console.error('styles.xml 解析错误');
      return {};
    }

    // 解析字体
    const fonts = [];
    const fontNodes = xmlDoc.getElementsByTagName('font');
    for (let i = 0; i < fontNodes.length; i++) {
      try {
        const font = {};
        const sz = fontNodes[i].getElementsByTagName('sz')[0];
        if (sz) { const val = parseFloat(sz.getAttribute('val')); if (!isNaN(val)) font.sz = val; }
        const name = fontNodes[i].getElementsByTagName('name')[0];
        if (name) font.name = name.getAttribute('val') || '';
        const bold = fontNodes[i].getElementsByTagName('b')[0];
        if (bold) font.bold = true;
        const italic = fontNodes[i].getElementsByTagName('i')[0];
        if (italic) font.italic = true;
        const underline = fontNodes[i].getElementsByTagName('u')[0];
        if (underline) font.underline = true;
        const strike = fontNodes[i].getElementsByTagName('strike')[0];
        if (strike) font.strike = true;
        const color = fontNodes[i].getElementsByTagName('color')[0];
        if (color) {
          font.color = {};
          const rgb = color.getAttribute('rgb');
          if (rgb && rgb.length >= 6) font.color.rgb = rgb;
          const theme = color.getAttribute('theme');
          if (theme) { const t = parseInt(theme); if (!isNaN(t)) font.color.theme = t; }
        }
        fonts.push(font);
      } catch (e) {
        console.warn('解析字体', i, '失败:', e);
      }
    }

    // 解析填充
    const fills = [];
    const fillNodes = xmlDoc.getElementsByTagName('fill');
    for (let i = 0; i < fillNodes.length; i++) {
      try {
        const fill = {};
        const patternFill = fillNodes[i].getElementsByTagName('patternFill')[0];
        if (patternFill) {
          fill.patternType = patternFill.getAttribute('patternType') || 'none';
          const fgColor = patternFill.getElementsByTagName('fgColor')[0];
          if (fgColor) {
            const rgb = fgColor.getAttribute('rgb');
            if (rgb && rgb.length >= 6) fill.fgColor = { rgb: rgb };
          }
        }
        fills.push(fill);
      } catch (e) {
        console.warn('解析填充', i, '失败:', e);
      }
    }

    // 解析边框
    const borders = [];
    const borderNodes = xmlDoc.getElementsByTagName('border');
    for (let i = 0; i < borderNodes.length; i++) {
      try {
        const border = {};
        const sides = ['left', 'right', 'top', 'bottom', 'diagonal'];
        sides.forEach(side => {
          const sideEl = borderNodes[i].getElementsByTagName(side)[0];
          if (sideEl) {
            border[side] = { style: sideEl.getAttribute('style') || 'none' };
            const colorEl = sideEl.getElementsByTagName('color')[0];
            if (colorEl) {
              const rgb = colorEl.getAttribute('rgb');
              if (rgb && rgb.length >= 6) border[side].color = { rgb: rgb };
            }
          }
        });
        borders.push(border);
      } catch (e) {
        console.warn('解析边框', i, '失败:', e);
      }
    }

    // 解析数字格式
    const numFmts = [];
    const numFmtNodes = xmlDoc.getElementsByTagName('numFmt');
    for (let i = 0; i < numFmtNodes.length; i++) {
      try {
        const numFmtId = parseInt(numFmtNodes[i].getAttribute('numFmtId'));
        if (!isNaN(numFmtId)) {
          numFmts.push({ numFmtId: numFmtId, formatCode: numFmtNodes[i].getAttribute('formatCode') || '' });
        }
      } catch (e) {
        console.warn('解析数字格式', i, '失败:', e);
      }
    }

    // 解析单元格样式 (cellXfs)
    const cellXfs = [];
    const xfNodes = xmlDoc.getElementsByTagName('cellXfs')[0];
    if (xfNodes) {
      const xfs = xfNodes.getElementsByTagName('xf');
      for (let i = 0; i < xfs.length; i++) {
        try {
          const xf = {};
          let numFmtId = parseInt(xfs[i].getAttribute('numFmtId'));
          xf.numFmtId = isNaN(numFmtId) ? 0 : numFmtId;
          let fontId = parseInt(xfs[i].getAttribute('fontId'));
          xf.fontId = isNaN(fontId) ? 0 : fontId;
          let fillId = parseInt(xfs[i].getAttribute('fillId'));
          xf.fillId = isNaN(fillId) ? 0 : fillId;
          let borderId = parseInt(xfs[i].getAttribute('borderId'));
          xf.borderId = isNaN(borderId) ? 0 : borderId;
          xf.applyNumberFormat = xfs[i].getAttribute('applyNumberFormat') === '1';
          xf.applyFont = xfs[i].getAttribute('applyFont') === '1';
          xf.applyFill = xfs[i].getAttribute('applyFill') === '1';
          xf.applyBorder = xfs[i].getAttribute('applyBorder') === '1';
          xf.applyAlignment = xfs[i].getAttribute('applyAlignment') === '1';

          const alignment = xfs[i].getElementsByTagName('alignment')[0];
          if (alignment) {
            xf.alignment = {};
            const horiz = alignment.getAttribute('horizontal');
            if (horiz) xf.alignment.horizontal = horiz;
            const vert = alignment.getAttribute('vertical');
            if (vert) xf.alignment.vertical = vert;
            if (alignment.getAttribute('wrapText') === '1') xf.alignment.wrapText = true;
            const indent = alignment.getAttribute('indent');
            if (indent) { const v = parseInt(indent); if (!isNaN(v)) xf.alignment.indent = v; }
            const textRotation = alignment.getAttribute('textRotation');
            if (textRotation) { const v = parseInt(textRotation); if (!isNaN(v)) xf.alignment.textRotation = v; }
          }
          cellXfs.push(xf);
        } catch (e) {
          console.warn('解析cellXfs', i, '失败:', e);
        }
      }
    }

    window._parsedStyles = {
      fonts: fonts, fills: fills, borders: borders, numFmts: numFmts, cellXfs: cellXfs
    };

    console.log('样式解析完成:', { fonts: fonts.length, fills: fills.length, borders: borders.length, cellXfs: cellXfs.length });

    return window._parsedStyles;

  } catch (e) {
    console.error('解析样式失败:', e);
    return {};
  }
}

// 根据样式索引构建完整的样式对象
function buildStyleFromIndex(styleIndex) {
  if (!window._parsedStyles || !window._parsedStyles.cellXfs) return {};
  if (styleIndex === undefined || styleIndex === null) return {};

  const xf = window._parsedStyles.cellXfs[styleIndex];
  if (!xf) return {};

  try {
    const style = {};

    // 字体样式
    if (xf.applyFont && xf.fontId !== undefined) {
      const font = window._parsedStyles.fonts[xf.fontId];
      if (font && typeof font === 'object' && Object.keys(font).length > 0) {
        style.font = {};
        if (font.sz && typeof font.sz === 'number') style.font.sz = font.sz;
        if (font.name && typeof font.name === 'string') style.font.name = font.name;
        if (font.bold === true) style.font.bold = true;
        if (font.italic === true) style.font.italic = true;
        if (font.underline === true) style.font.underline = true;
        if (font.strike === true) style.font.strike = true;
        if (font.color && typeof font.color === 'object' && font.color.rgb) {
          style.font.color = { rgb: font.color.rgb };
        }
      }
    }

    // 填充样式
    if (xf.applyFill && xf.fillId !== undefined) {
      const fill = window._parsedStyles.fills[xf.fillId];
      if (fill && fill.fgColor && fill.fgColor.rgb && fill.fgColor.rgb.length >= 6) {
        style.fill = { fgColor: { rgb: fill.fgColor.rgb } };
      }
    }

    // 边框样式
    if (xf.applyBorder && xf.borderId !== undefined) {
      const border = window._parsedStyles.borders[xf.borderId];
      if (border && typeof border === 'object' && Object.keys(border).length > 0) {
        style.border = {};
        ['left', 'right', 'top', 'bottom'].forEach(side => {
          if (border[side] && border[side].style && border[side].style !== 'none') {
            style.border[side] = { style: border[side].style };
            if (border[side].color && border[side].color.rgb && border[side].color.rgb.length >= 6) {
              style.border[side].color = { rgb: border[side].color.rgb };
            }
          }
        });
        if (Object.keys(style.border).length === 0) delete style.border;
      }
    }

    // 对齐方式
    if (xf.applyAlignment && xf.alignment && typeof xf.alignment === 'object') {
      style.alignment = {};
      if (xf.alignment.horizontal) style.alignment.horizontal = xf.alignment.horizontal;
      if (xf.alignment.vertical) style.alignment.vertical = xf.alignment.vertical;
      if (xf.alignment.wrapText === true) style.alignment.wrapText = true;
      if (xf.alignment.indent !== undefined && typeof xf.alignment.indent === 'number') style.alignment.indent = xf.alignment.indent;
      if (xf.alignment.textRotation !== undefined && typeof xf.alignment.textRotation === 'number') style.alignment.textRotation = xf.alignment.textRotation;
      if (Object.keys(style.alignment).length === 0) delete style.alignment;
    }

    // 数字格式
    if (xf.applyNumberFormat && xf.numFmtId && xf.numFmtId > 0) {
      style.numFmt = xf.numFmtId;
    }

    return style;
  } catch (e) {
    console.warn('构建样式对象失败:', styleIndex, e);
    return {};
  }
}

// 从 sheet XML 中直接提取单元格的样式索引
// 这是关键修复：xlsx-js-style 不填充 cell.s，所以我们从原始 XML 读取
async function extractCellStyleIndicesFromXML(fileData, workbook) {
  const cellStyleIndices = {};
  try {
    if (!fileData || !workbook || !workbook.SheetNames) return cellStyleIndices;

    const zip = await JSZip.loadAsync(fileData);
    if (!zip) return cellStyleIndices;

    for (let sheetIdx = 0; sheetIdx < workbook.SheetNames.length; sheetIdx++) {
      const sheetName = workbook.SheetNames[sheetIdx];
      cellStyleIndices[sheetName] = {};
      const sheetFile = zip.file('xl/worksheets/sheet' + (sheetIdx + 1) + '.xml');
      if (!sheetFile) continue;

      const sheetText = await sheetFile.async('text');
      if (!sheetText) continue;

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(sheetText, 'text/xml');
      if (xmlDoc.documentElement && xmlDoc.documentElement.tagName === 'parsererror') continue;

      const rows = xmlDoc.getElementsByTagName('row');
      for (let r = 0; r < rows.length; r++) {
        const rowAttr = rows[r].getAttribute('r');
        const rowIdx = rowAttr ? parseInt(rowAttr) - 1 : r;
        if (isNaN(rowIdx)) continue;
        const cells = rows[r].getElementsByTagName('c');
        for (let c = 0; c < cells.length; c++) {
          const cell = cells[c];
          const rAttr = cell.getAttribute('r');
          let colIdx = 0;
          if (rAttr) {
            const colMatch = rAttr.match(/[A-Z]+/);
            if (colMatch) {
              const colStr = colMatch[0];
              for (let i = 0; i < colStr.length; i++) {
                colIdx = colIdx * 26 + (colStr.charCodeAt(i) - 64);
              }
              colIdx -= 1;
            } else {
              colIdx = c;
            }
          } else {
            colIdx = c;
          }
          const sAttr = cell.getAttribute('s');
          if (sAttr !== null && sAttr !== undefined) {
            const styleIdx = parseInt(sAttr);
            if (isNaN(styleIdx)) continue;
            if (!cellStyleIndices[sheetName][rowIdx]) {
              cellStyleIndices[sheetName][rowIdx] = {};
            }
            cellStyleIndices[sheetName][rowIdx][colIdx] = styleIdx;
          }
        }
      }
    }
  } catch (e) {
    console.error('从XML提取样式索引失败:', e);
  }
  return cellStyleIndices;
}

// 主函数：从原始数据和 workbook 中提取完整样式
async function extractAllStyles(fileData, workbook) {
  // 1. 解析 styles.xml
  const parsed = await extractStylesFromXlsx(fileData);

  // 2. 从 sheet XML 提取单元格的样式索引（关键修复）
  const indices = await extractCellStyleIndicesFromXML(fileData, workbook);

  // 3. 构建每个单元格的完整样式对象
  const sheetStyles = {};
  workbook.SheetNames.forEach(sheetName => {
    sheetStyles[sheetName] = {};
    const indicesForSheet = indices[sheetName] || {};
    Object.keys(indicesForSheet).forEach(rowStr => {
      const row = parseInt(rowStr);
      sheetStyles[sheetName][row] = {};
      Object.keys(indicesForSheet[row]).forEach(colStr => {
        const col = parseInt(colStr);
        const styleIdx = indicesForSheet[row][col];
        sheetStyles[sheetName][row][col] = buildStyleFromIndex(styleIdx);
      });
    });
  });

  return sheetStyles;
}
