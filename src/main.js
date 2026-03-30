const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const http = require('http');
const { URL } = require('url');
const MarkdownIt = require('markdown-it');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType, ImageRun, ShadingType, VerticalAlign, PageOrientation } = require('docx');
const httpsAgent = new https.Agent({ keepAlive: true });

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 650,
    height: 580,
    resizable: true,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js'),
      sandbox: true,
    }
  });
  mainWindow.loadFile('src/index.html');
  mainWindow.setMenu(null);
}

async function downloadImage(imageUrl) {
  return new Promise((resolve, reject) => {
    try {
      const parsedUrl = new URL(imageUrl);
      const protocol = parsedUrl.protocol === 'https:' ? https : http;
      const requestOptions = {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
          'Referer': parsedUrl.origin,
        },
        timeout: 15000,
      };
      if (parsedUrl.protocol === 'https:') {
        requestOptions.agent = httpsAgent;
      }
      const req = protocol.get(imageUrl, requestOptions, (res) => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          downloadImage(res.headers.location).then(resolve).catch(reject);
          return;
        }
        if (res.statusCode !== 200) {
          resolve(null);
          return;
        }
        const chunks = [];
        res.on('data', (chunk) => chunks.push(chunk));
        res.on('end', () => {
          const buffer = Buffer.concat(chunks);
          resolve(buffer);
        });
        res.on('error', () => resolve(null));
      });
      req.on('error', () => resolve(null));
      req.on('timeout', () => { req.destroy(); resolve(null); });
    } catch (e) {
      resolve(null);
    }
  });
}

function parseInlineStyles(tokens) {
  const runs = [];
  let text = '';
  let inCode = false;
  let codeContent = '';

  for (const token of tokens) {
    if (token.type === 'inline') {
      for (const child of token.children || []) {
        const content = child.content || '';
        if (child.type === 'text') {
          if (inCode) {
            codeContent += content;
          } else {
            text += content;
          }
        } else if (child.type === 'code_inline') {
          if (!inCode && text) {
            runs.push(new TextRun({ text, bold: false, italics: false }));
            text = '';
          }
          inCode = !inCode;
          if (!inCode && codeContent) {
            runs.push(new TextRun({ text: codeContent, font: 'Courier New', size: 18, shading: { fill: 'F0F0F0' } }));
            codeContent = '';
          }
        }
      }
    }
  }
  if (text) {
    runs.push(new TextRun({ text, bold: false, italics: false }));
  }
  if (inCode && codeContent) {
    runs.push(new TextRun({ text: codeContent, font: 'Courier New', size: 18, shading: { fill: 'F0F0F0' } }));
  }
  return runs;
}

function mdInlineToRuns(text) {
  const runs = [];
  let remaining = text;

  while (remaining.length > 0) {
    // Bold + Italic ***text***
    let m = remaining.match(/^\*\*\*(.+?)\*\*\*/);
    if (m) {
      runs.push(new TextRun({ text: m[1], bold: true, italics: true }));
      remaining = remaining.substring(m[0].length);
      continue;
    }
    // Bold **text**
    m = remaining.match(/^\*\*(.+?)\*\*/);
    if (m) {
      runs.push(new TextRun({ text: m[1], bold: true }));
      remaining = remaining.substring(m[0].length);
      continue;
    }
    // Italic *text*
    m = remaining.match(/^\*(.+?)\*/);
    if (m) {
      runs.push(new TextRun({ text: m[1], italics: true }));
      remaining = remaining.substring(m[0].length);
      continue;
    }
    // Strikethrough ~~text~~
    m = remaining.match(/^~~(.+?)~~/);
    if (m) {
      runs.push(new TextRun({ text: m[1], strike: true }));
      remaining = remaining.substring(m[0].length);
      continue;
    }
    // Inline code `text`
    m = remaining.match(/^`([^`]+)`/);
    if (m) {
      runs.push(new TextRun({ text: m[1], font: 'Courier New', size: 18, shading: { fill: 'F0F0F0' } }));
      remaining = remaining.substring(m[0].length);
      continue;
    }
    // Any other character
    const nextSpecial = remaining.search(/\*\*|^\*|~~|`/);
    if (nextSpecial === -1) {
      runs.push(new TextRun({ text: remaining }));
      break;
    } else if (nextSpecial === 0) {
      runs.push(new TextRun({ text: remaining.charAt(0) }));
      remaining = remaining.substring(1);
    } else {
      runs.push(new TextRun({ text: remaining.substring(0, nextSpecial) }));
      remaining = remaining.substring(nextSpecial);
    }
  }
  return runs.length > 0 ? runs : [new TextRun({ text: '' })];
}

async function parseMarkdownToDocx(content, baseDir) {
  const md = new MarkdownIt({
    html: false,
    linkify: true,
    typographer: true,
  });
  const tokens = md.parse(content, {});

  const docChildren = [];
  let i = 0;
  const len = tokens.length;

  const inBlockquote = () => {
    let depth = 0;
    for (let j = i - 1; j >= 0; j--) {
      if (tokens[j].type === 'blockquote_open') depth++;
      else if (tokens[j].type === 'blockquote_close') depth--;
      if (depth > 0) return true;
    }
    return false;
  };

  while (i < len) {
    const token = tokens[i];

    if (token.type === 'heading_open') {
      const level = parseInt(token.tag.substring(1));
      i++;
      const inlineTokens = [];
      while (i < len && tokens[i].type !== 'heading_close') {
        if (tokens[i].type === 'inline') inlineTokens.push(...(tokens[i].children || []));
        else if (tokens[i].type === 'softbreak') inlineTokens.push({ type: 'text', content: ' ' });
        i++;
      }
      const rawText = inlineTokens.map(t => t.content || '').join('');
      let headingLevel;
      switch (level) {
        case 1: headingLevel = HeadingLevel.HEADING_1; break;
        case 2: headingLevel = HeadingLevel.HEADING_2; break;
        case 3: headingLevel = HeadingLevel.HEADING_3; break;
        case 4: headingLevel = HeadingLevel.HEADING_4; break;
        case 5: headingLevel = HeadingLevel.HEADING_5; break;
        default: headingLevel = HeadingLevel.HEADING_6;
      }
      const runs = mdInlineToRuns(rawText);
      docChildren.push(new Paragraph({
        children: runs,
        heading: headingLevel,
        spacing: { before: 240, after: 120 },
        border: level <= 2 ? {
          bottom: { color: 'CCCCCC', space: 1, style: 'single', size: 4 },
        } : undefined,
      }));
      i++;
    } else if (token.type === 'paragraph_open') {
      i++;
      const inlineTokens = [];
      while (i < len && tokens[i].type !== 'paragraph_close') {
        if (tokens[i].type === 'inline') inlineTokens.push(...(tokens[i].children || []));
        else if (tokens[i].type === 'softbreak') inlineTokens.push({ type: 'text', content: ' ' });
        else if (tokens[i].type === 'hardbreak') inlineTokens.push({ type: 'text', content: '\n' });
        i++;
      }
      const rawText = inlineTokens.map(t => t.content || '').join('');
      if (rawText.trim()) {
        const runs = mdInlineToRuns(rawText);
        docChildren.push(new Paragraph({
          children: runs,
          spacing: { after: 200 },
          indent: inBlockquote() ? { left: 720 } : undefined,
        }));
      }
      i++;
    } else if (token.type === 'bullet_list_open') {
      let j = i + 1;
      let listEnded = false;
      while (j < len && !listEnded) {
        if (tokens[j].type === 'bullet_list_close') { listEnded = true; break; }
        if (tokens[j].type === 'list_item_open') {
          j++;
          let itemText = '';
          let subList = false;
          while (j < len && tokens[j].type !== 'list_item_close') {
            if (tokens[j].type === 'inline') {
              itemText += (tokens[j].content || '');
            } else if (tokens[j].type === 'bullet_list_open' || tokens[j].type === 'ordered_list_open') {
              subList = true;
            }
            j++;
          }
          if (!subList) {
            const runs = mdInlineToRuns(itemText.trim());
            docChildren.push(new Paragraph({
              children: runs,
              bullet: { level: 0 },
              spacing: { after: 80 },
            }));
          }
        }
        j++;
      }
      i = j + 1;
    } else if (token.type === 'ordered_list_open') {
      let j = i + 1;
      let num = 1;
      let listEnded = false;
      while (j < len && !listEnded) {
        if (tokens[j].type === 'ordered_list_close') { listEnded = true; break; }
        if (tokens[j].type === 'list_item_open') {
          j++;
          let itemText = '';
          let subList = false;
          while (j < len && tokens[j].type !== 'list_item_close') {
            if (tokens[j].type === 'inline') itemText += (tokens[j].content || '');
            else if (tokens[j].type === 'bullet_list_open' || tokens[j].type === 'ordered_list_open') subList = true;
            j++;
          }
          if (!subList) {
            const runs = mdInlineToRuns(itemText.trim());
            docChildren.push(new Paragraph({
              children: runs,
              numbering: { reference: 'default-numbering', level: 0 },
              spacing: { after: 80 },
            }));
          }
          num++;
        }
        j++;
      }
      i = j + 1;
    } else if (token.type === 'blockquote_open') {
      i++;
      let quoteText = '';
      let depth = 1;
      while (i < len) {
        if (tokens[i].type === 'blockquote_open') { depth++; i++; }
        else if (tokens[i].type === 'blockquote_close') {
          depth--;
          if (depth === 0) { i++; break; }
          i++;
        } else {
          if (tokens[i].type === 'inline') quoteText += (tokens[i].content || '');
          else if (tokens[i].type === 'softbreak') quoteText += ' ';
          else if (tokens[i].type === 'paragraph_close') { i++; continue; }
          i++;
        }
      }
      if (quoteText.trim()) {
        const runs = mdInlineToRuns(quoteText.trim());
        docChildren.push(new Paragraph({
          children: runs,
          spacing: { after: 120 },
          indent: { left: 720 },
          border: {
            left: { color: '888888', space: 8, style: 'single', size: 12 },
          },
        }));
      }
    } else if (token.type === 'hr') {
      docChildren.push(new Paragraph({
        children: [new TextRun({ text: '' })],
        border: {
          bottom: { color: 'CCCCCC', space: 1, style: 'single', size: 6 },
        },
        spacing: { before: 200, after: 200 },
      }));
      i++;
    } else if (token.type === 'code_block') {
      const code = token.content || '';
      const lang = token.info || '';
      docChildren.push(new Paragraph({
        children: [
          new TextRun({ text: lang ? `\`\`\`${lang}\n` + code + '```' : code, font: 'Courier New', size: 18, shading: { fill: 'F5F5F5' } }),
        ],
        spacing: { before: 160, after: 160 },
      }));
      i++;
    } else if (token.type === 'fence') {
      const code = token.content || '';
      const lang = token.info || '';
      const lines = code.split('\n');
      const paragraphs = [];
      if (lang) {
        paragraphs.push(new Paragraph({
          children: [new TextRun({ text: `Language: ${lang}`, font: 'Courier New', size: 16, color: '888888', italics: true })],
          spacing: { after: 40 },
        }));
      }
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: code || '(empty code block)', font: 'Courier New', size: 18, shading: { fill: 'F5F5F5' } })],
        spacing: { before: 120, after: 120 },
      }));
      docChildren.push(...paragraphs);
      i++;
    } else if (token.type === 'table_open') {
      const tableRows = [];
      let j = i + 1;
      let rowIndex = 0;
      const alignments = [];

      while (j < len && tokens[j].type !== 'table_close') {
        if (tokens[j].type === 'tr_open') {
          j++;
          const cells = [];
          let cellIndex = 0;
          while (j < len && tokens[j].type !== 'tr_close') {
            if (tokens[j].type === 'th_open' || tokens[j].type === 'td_open') {
              const isHeader = tokens[j].type === 'th_open';
              let align = AlignmentType.LEFT;
              if (tokens[j].tag === 'th' || tokens[j].tag === 'td') {
                const alignAttr = tokens[j].info || tokens[j].attrGet && tokens[j].attrGet('align');
                if (alignAttr === 'center') align = AlignmentType.CENTER;
                else if (alignAttr === 'right') align = AlignmentType.RIGHT;
                if (isHeader) alignments[cellIndex] = align;
              }
              j++;
              let cellText = '';
              while (j < len && tokens[j].type !== 'th_close' && tokens[j].type !== 'td_close') {
                if (tokens[j].type === 'inline') {
                  const children = tokens[j].children || [];
                  cellText += children.map(c => c.content || '').join('');
                }
                j++;
              }
              cells.push(new TableCell({
                children: [new Paragraph({
                  children: mdInlineToRuns(cellText),
                  alignment: align,
                })],
                shading: { fill: isHeader ? 'D9D9D9' : undefined, type: ShadingType.CLEAR, color: 'auto' },
                verticalAlign: VerticalAlign.CENTER,
                width: { size: null, type: WidthType.AUTO },
              }));
              cellIndex++;
            }
            j++;
          }
          tableRows.push(new TableRow({
            children: cells,
            tableHeader: rowIndex === 0,
          }));
          rowIndex++;
        }
        j++;
      }
      if (tableRows.length > 0) {
        docChildren.push(new Table({
          rows: tableRows,
          width: { size: 100, type: WidthType.PERCENTAGE },
        }));
        docChildren.push(new Paragraph({ spacing: { after: 200 } }));
      }
      i = j + 1;
    } else if (token.type === 'image') {
      const src = token.attrGet('src') || '';
      const alt = token.attrGet('alt') || '图片';
      const altText = alt.replace(/!\[|\]|\*/g, '');

      docChildren.push(new Paragraph({
        children: [new TextRun({ text: `[图片: ${altText}]`, italics: true, color: '666666', size: 20 })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
      }));

      if (src.startsWith('http://') || src.startsWith('https://')) {
        try {
          const imgBuffer = await downloadImage(src);
          if (imgBuffer) {
            const ext = src.split('.').pop().toLowerCase().split('?')[0];
            const mimeMap = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' };
            const mime = mimeMap[ext] || 'image/png';
            const base64 = imgBuffer.toString('base64');
            docChildren.push(new Paragraph({
              children: [
                new ImageRun({
                  data: imgBuffer,
                  transformation: { width: 400, height: 300 },
                  type: 'png',
                  mimeType: mime,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 120 },
            }));
          } else {
            docChildren.push(new Paragraph({
              children: [new TextRun({ text: `URL: ${src}`, font: 'Calibri', size: 18, color: '0563C1' })],
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
            }));
          }
        } catch (e) {
          docChildren.push(new Paragraph({
            children: [new TextRun({ text: `URL: ${src}`, font: 'Calibri', size: 18, color: '0563C1' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }));
        }
      } else {
        const absolutePath = path.isAbsolute(src) ? src : path.join(baseDir || '', src);
        if (fs.existsSync(absolutePath)) {
          try {
            const imgBuffer = fs.readFileSync(absolutePath);
            const ext = path.extname(absolutePath).substring(1).toLowerCase();
            const mimeMap = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' };
            const mime = mimeMap[ext] || 'image/png';
            docChildren.push(new Paragraph({
              children: [
                new ImageRun({
                  data: imgBuffer,
                  transformation: { width: 400, height: 300 },
                  type: 'png',
                  mimeType: mime,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { after: 120 },
            }));
          } catch (e) {
            docChildren.push(new Paragraph({
              children: [new TextRun({ text: `[本地图片: ${src}（读取失败）]`, font: 'Calibri', size: 18, color: 'C00000' })],
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
            }));
          }
        } else {
          docChildren.push(new Paragraph({
            children: [new TextRun({ text: `[本地图片: ${src}（文件不存在）]`, font: 'Calibri', size: 18, color: 'C00000' })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }));
        }
      }
      i++;
    } else {
      i++;
    }
  }

  const doc = new Document({
    numbering: {
      config: [{
        reference: 'default-numbering',
        levels: [{
          level: 0,
          format: 'decimal',
          text: '%1.',
          alignment: AlignmentType.LEFT,
          style: {
            paragraph: { indent: { left: 720, hanging: 360 } },
          },
        }],
      }],
    },
    sections: [{
      properties: {
        page: {
          orientation: PageOrientation.A4,
        },
      },
      children: docChildren,
    }],
  });
  return doc;
}

ipcMain.handle('convert-md-to-docx', async (event, mdContent, filePath) => {
  try {
    const baseDir = filePath ? path.dirname(filePath) : '';
    const doc = await parseMarkdownToDocx(mdContent, baseDir);
    const buffer = await Packer.toBuffer(doc);
    return { success: true, buffer: buffer.toString('base64') };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('open-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Markdown', extensions: ['md', 'markdown', 'txt'] }],
  });
  if (!result.canceled && result.filePaths.length > 0) {
    const filePath = result.filePaths[0];
    const content = fs.readFileSync(filePath, 'utf-8');
    return { filePath, content };
  }
  return null;
});

ipcMain.handle('save-file', async (event, base64Data, originalName) => {
  const basePath = originalName.replace(/\.[^.]+$/, '.docx');
  let savePath = basePath;
  let counter = 1;
  const buffer = Buffer.from(base64Data, 'base64');

  while (true) {
    try {
      fs.writeFileSync(savePath, buffer);
      return { success: true, filePath: savePath };
    } catch (error) {
      if (error.code === 'EBUSY' || error.code === 'ENOENT') {
        const dir = path.dirname(savePath);
        const basename = path.basename(savePath, path.extname(savePath));
        const ext = path.extname(savePath);
        savePath = path.join(dir, `${basename} (${counter})${ext}`);
        counter++;
        if (counter > 20) {
          return { success: false, error: '文件被占用，请关闭后重试' };
        }
      } else {
        return { success: false, error: error.message };
      }
    }
  }
});

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  app.quit();
});
