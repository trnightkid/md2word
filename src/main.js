const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const MarkdownIt = require('markdown-it');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ImageRun } = require('docx');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 650,
    height: 550,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });
  mainWindow.loadFile('src/index.html');
  mainWindow.setMenu(null);
}

function parseMarkdownToDocx(content) {
  const md = new MarkdownIt();
  const tokens = md.parse(content, {});
  
  const docChildren = [];
  let imageCount = 0;
  
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    
    if (token.type === 'heading_open' && tokens[i + 1] && tokens[i + 1].type === 'inline') {
      const level = parseInt(token.tag.substring(1));
      const text = tokens[i + 1].content;
      
      let headingLevel;
      switch(level) {
        case 1: headingLevel = HeadingLevel.HEADING_1; break;
        case 2: headingLevel = HeadingLevel.HEADING_2; break;
        case 3: headingLevel = HeadingLevel.HEADING_3; break;
        default: headingLevel = HeadingLevel.HEADING_4;
      }
      
      docChildren.push(new Paragraph({
        text: text,
        heading: headingLevel,
        spacing: { after: 200 }
      }));
      i += 2;
    } else if (token.type === 'paragraph_open' && tokens[i + 1] && tokens[i + 1].type === 'inline') {
      const text = tokens[i + 1].content;
      docChildren.push(new Paragraph({
        children: [new TextRun(text)],
        spacing: { after: 200 }
      }));
      i += 2;
    } else if (token.type === 'bullet_list_open') {
      let j = i + 1;
      while (j < tokens.length && tokens[j].type !== 'bullet_list_close') {
        if (tokens[j].type === 'list_item_open') {
          j++;
          if (tokens[j] && tokens[j].type === 'inline') {
            docChildren.push(new Paragraph({
              text: '• ' + tokens[j].content,
              indent: { left: 720 },
              spacing: { after: 100 }
            }));
          }
        }
        j++;
      }
      i = j;
    } else if (token.type === 'ordered_list_open') {
      let j = i + 1;
      let num = 1;
      while (j < tokens.length && tokens[j].type !== 'ordered_list_close') {
        if (tokens[j].type === 'list_item_open') {
          j++;
          if (tokens[j] && tokens[j].type === 'inline') {
            docChildren.push(new Paragraph({
              text: num + '. ' + tokens[j].content,
              indent: { left: 720 },
              spacing: { after: 100 }
            }));
            num++;
          }
        }
        j++;
      }
      i = j;
    } else if (token.type === 'hr') {
      docChildren.push(new Paragraph({
        text: '─────────────────────────────────────',
        spacing: { after: 200 }
      }));
    } else if (token.type === 'table_open') {
      const tableRows = [];
      let j = i + 1;
      while (j < tokens.length && tokens[j].type !== 'table_close') {
        if (tokens[j].type === 'tr_open') {
          const cells = [];
          j++;
          while (j < tokens.length && tokens[j].type !== 'tr_close') {
            if (tokens[j].type === 'th_open' || tokens[j].type === 'td_open') {
              const isHeader = tokens[j].type === 'th_open';
              j++;
              let cellText = '';
              if (tokens[j] && tokens[j].type === 'inline') {
                if (tokens[j].children && tokens[j].children.length > 0) {
                  cellText = tokens[j].children.map(c => c.content).join('');
                } else {
                  cellText = tokens[j].content || '';
                }
              }
              cells.push(new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ text: cellText, bold: isHeader })],
                })],
                shading: { fill: isHeader ? 'E7E6E6' : undefined },
              }));
            }
            j++;
          }
          tableRows.push(new TableRow({ children: cells }));
        }
        j++;
      }
      if (tableRows.length > 0) {
        docChildren.push(new Table({
          rows: tableRows,
          width: { size: 100, type: WidthType.PERCENTAGE }
        }));
        docChildren.push(new Paragraph({ spacing: { after: 200 } }));
      }
      i = j;
    } else if (token.type === 'image') {
      const src = token.attrGet('src');
      const alt = token.attrGet('alt') || '图片';
      
      if (src.startsWith('http://') || src.startsWith('https://')) {
        docChildren.push(new Paragraph({
          children: [new TextRun({ text: '[图片: ' + alt + ']', italics: true, color: '666666' })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 }
        }));
        docChildren.push(new Paragraph({
          children: [new TextRun({ text: 'URL: ' + src, font: 'Calibri', size: 18, color: '0000FF' })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 }
        }));
      } else {
        docChildren.push(new Paragraph({
          children: [new TextRun({ text: '[图片: ' + alt + ']', italics: true, color: '666666' })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 }
        }));
        docChildren.push(new Paragraph({
          children: [new TextRun({ text: '图片路径: ' + src, font: 'Calibri', size: 18 })],
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 }
        }));
      }
      imageCount++;
      i++;
    }
  }
  
  return new Document({
    sections: [{
      properties: {},
      children: docChildren
    }]
  });
}

ipcMain.handle('convert-md-to-docx', async (event, mdContent) => {
  try {
    const doc = parseMarkdownToDocx(mdContent);
    const buffer = await Packer.toBuffer(doc);
    return { success: true, buffer: buffer.toString('base64') };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('open-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Markdown', extensions: ['md', 'markdown', 'txt'] }]
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
        const ext = savePath.lastIndexOf('.');
        const name = ext > 0 ? savePath.substring(0, ext) : savePath;
        const extName = ext > 0 ? savePath.substring(ext) : '.docx';
        const dir = savePath.substring(0, savePath.lastIndexOf('\\') + 1);
        savePath = dir + name + ' (' + counter + ')' + extName;
        counter++;
        if (counter > 10) {
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
