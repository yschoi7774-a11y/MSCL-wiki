require('dotenv').config();
const express = require('express');
const fs = require('fs');
const path = require('path');
const { marked } = require('marked');
const Anthropic = require('@anthropic-ai/sdk');
const session = require('express-session');
const PptxGenJS = require('pptxgenjs');
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType, ShadingType } = require('docx');

const anthropic = process.env.ANTHROPIC_API_KEY
  ? new Anthropic()
  : null;

const app = express();
const PORT = process.env.PORT || 3000;
const WIKI_DIR = path.join(__dirname, 'wiki');

// ── 인증 ──────────────────────────────────────────
app.use(session({ secret: process.env.SESSION_SECRET || 'mscl-2026', resave: false, saveUninitialized: false, cookie: { maxAge: 7*24*60*60*1000 } }));
app.use(express.urlencoded({ extended: false }));
app.use(express.json());

const LOGIN_HTML = `<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>로그인</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:-apple-system,sans-serif;background:#f0f2f5;display:flex;align-items:center;justify-content:center;height:100vh}.box{background:#fff;border-radius:16px;padding:44px 40px;width:320px;box-shadow:0 4px 24px rgba(0,0,0,0.1)}h1{font-size:17px;color:#1a1a2e;margin-bottom:28px;text-align:center;font-weight:700}input{width:100%;padding:11px 14px;border:1.5px solid #e0e0e0;border-radius:8px;font-size:14px;margin-bottom:12px;outline:none;color:#222}input:focus{border-color:#4a6fa5}button{width:100%;padding:12px;background:#1a1a2e;color:#fff;border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;margin-top:4px}button:hover{background:#2a2a4e}.err{color:#e53935;font-size:13px;text-align:center;margin-top:14px}</style></head><body><div class="box"><h1>🔐 MSCL 지식 위키</h1><form method="POST" action="/login"><input type="text" name="id" placeholder="아이디" autofocus autocomplete="username"><input type="password" name="password" placeholder="비밀번호" autocomplete="current-password"><button type="submit">로그인</button>__ERR__</form></div></body></html>`;

function requireAuth(req, res, next) {
  if (req.path === '/login' || req.session.loggedIn) return next();
  res.redirect('/login');
}

app.get('/login', (req, res) => {
  if (req.session.loggedIn) return res.redirect('/');
  res.send(LOGIN_HTML.replace('__ERR__', ''));
});
app.post('/login', (req, res) => {
  if (req.body.id === 'mothersmile' && req.body.password === '0544') {
    req.session.loggedIn = true; res.redirect('/');
  } else {
    res.send(LOGIN_HTML.replace('__ERR__', '<p class="err">아이디 또는 비밀번호가 틀렸습니다.</p>'));
  }
});
app.get('/logout', (req, res) => { req.session.destroy(); res.redirect('/login'); });

app.use(requireAuth);
// ─────────────────────────────────────────────────

app.use(express.static(path.join(__dirname, 'public')));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

function processWikiLinks(html) {
  html = html.replace(/\[\[([^\]|]+)\|([^\]]+)\]\]/g, (_, page, display) => {
    return `<a href="/wiki/${encodeURIComponent(page)}" class="wikilink">${display}</a>`;
  });
  html = html.replace(/\[\[([^\]]+)\]\]/g, (_, page) => {
    return `<a href="/wiki/${encodeURIComponent(page)}" class="wikilink">${page}</a>`;
  });
  return html;
}

function processCallouts(markdown) {
  markdown = markdown.replace(
    /^> \[!(\w+)\]-\s*(.*)\n((?:^>.*\n?)*)/gm,
    (_, type, title, body) => {
      const content = body.replace(/^> ?/gm, '').trim();
      return `<details class="callout callout-${type.toLowerCase()}"><summary><span class="callout-type">${type.toUpperCase()}</span>${title ? ' ' + title : ''}</summary><div class="callout-content">\n\n${content}\n\n</div></details>\n`;
    }
  );
  markdown = markdown.replace(
    /^> \[!(\w+)\]\s*(.*)\n((?:^>.*\n?)*)/gm,
    (_, type, title, body) => {
      const content = body.replace(/^> ?/gm, '').trim();
      return `<div class="callout callout-${type.toLowerCase()}"><div class="callout-title"><span class="callout-type">${type.toUpperCase()}</span>${title ? ' ' + title : ''}</div><div class="callout-content">\n\n${content}\n\n</div></div>\n`;
    }
  );
  return markdown;
}

function parseSidebar() {
  const indexPath = path.join(WIKI_DIR, 'index.md');
  if (!fs.existsSync(indexPath)) return [];
  const content = fs.readFileSync(indexPath, 'utf-8');
  const categories = [];
  let current = null;
  for (const line of content.split('\n')) {
    const h2 = line.match(/^## (.+)/);
    if (h2) { current = { name: h2[1].trim(), pages: [] }; categories.push(current); continue; }
    const link = line.match(/^- \[\[([^\]|]+)(?:\|[^\]]+)?\]\]/);
    if (link && current) {
      const pageName = link[1].trim();
      const dashMatch = line.match(/\]\] — (.+)/);
      current.pages.push({ name: pageName, desc: dashMatch ? dashMatch[1] : '' });
    }
  }
  return categories;
}

function renderPage(pageName) {
  const filePath = path.join(WIKI_DIR, `${pageName}.md`);
  if (!fs.existsSync(filePath)) return null;
  let content = fs.readFileSync(filePath, 'utf-8');
  content = processCallouts(content);
  let html = marked(content);
  html = processWikiLinks(html);
  return html;
}

// 위키 파일에서 [[링크]] 추출
function extractLinks(content) {
  const links = [];
  const re = /\[\[([^\]|]+)(?:\|[^\]]+)?\]\]/g;
  let m;
  while ((m = re.exec(content)) !== null) links.push(m[1].trim());
  return [...new Set(links)];
}

// 그래프 데이터 빌드
function buildGraph() {
  const sidebar = parseSidebar();
  const categoryMap = {};
  sidebar.forEach(cat => cat.pages.forEach(p => { categoryMap[p.name] = cat.name; }));

  const newPagesPath = path.join(__dirname, 'new-pages.json');
  const newSet = new Set(fs.existsSync(newPagesPath)
    ? JSON.parse(fs.readFileSync(newPagesPath, 'utf-8'))
    : []);

  const files = fs.readdirSync(WIKI_DIR).filter(f => f.endsWith('.md') && f !== 'log.md');
  const nodes = [];
  const links = [];
  const pageSet = new Set(files.map(f => f.replace('.md', '')));

  files.forEach(file => {
    const name = file.replace('.md', '');
    const content = fs.readFileSync(path.join(WIKI_DIR, file), 'utf-8');
    nodes.push({ id: name, category: categoryMap[name] || '기타', isNew: newSet.has(name) });
    extractLinks(content).forEach(target => {
      if (pageSet.has(target) && target !== name) {
        links.push({ source: name, target });
      }
    });
  });

  return { nodes, links };
}

// 위키 페이지에서 다운로드 링크 추출 ([[raw/파일명|라벨]] 형식)
function extractDownloads(raw) {
  const downloads = [];
  const re = /\[\[raw\/([^\]|]+?)(?:\|([^\]]+))?\]\]/g;
  let m;
  while ((m = re.exec(raw)) !== null) {
    downloads.push({ file: m[1].trim(), label: (m[2] || m[1]).replace(/^⬇\s*/, '').trim() });
  }
  return downloads;
}

// 검색 인덱스 빌드
function buildIndex() {
  const files = fs.readdirSync(WIKI_DIR).filter(f => f.endsWith('.md') && f !== 'log.md');
  return files.map(file => {
    const name = file.replace('.md', '');
    const raw = fs.readFileSync(path.join(WIKI_DIR, file), 'utf-8');
    // 마크다운 기호 제거
    const text = raw.replace(/[#*`>\[\]_~]/g, ' ').replace(/\s+/g, ' ');
    return { name, text, raw };
  });
}

function search(query) {
  const index = buildIndex();
  const tokens = query.toLowerCase().split(/\s+/).filter(Boolean);

  return index
    .map(page => {
      const haystack = page.text.toLowerCase();
      const titleHaystack = page.name.toLowerCase();
      let score = 0;
      tokens.forEach(t => {
        const titleCount = (titleHaystack.match(new RegExp(t, 'g')) || []).length;
        const bodyCount = (haystack.match(new RegExp(t, 'g')) || []).length;
        score += titleCount * 5 + bodyCount;
      });
      if (score === 0) return null;

      // 스니펫: 첫 번째 토큰 주변 100자
      let snippet = '';
      for (const t of tokens) {
        const idx = haystack.indexOf(t);
        if (idx !== -1) {
          const start = Math.max(0, idx - 60);
          const end = Math.min(page.text.length, idx + 100);
          snippet = (start > 0 ? '…' : '') + page.text.slice(start, end).trim() + '…';
          break;
        }
      }
      // 토큰 하이라이트
      tokens.forEach(t => {
        snippet = snippet.replace(new RegExp(`(${t})`, 'gi'), '<mark>$1</mark>');
      });

      return { name: page.name, score, snippet, downloads: extractDownloads(page.raw) };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);
}

// AI 질문 답변
async function askAI(question) {
  if (!anthropic) return null;
  const files = fs.readdirSync(WIKI_DIR)
    .filter(f => f.endsWith('.md') && f !== 'log.md' && f !== 'index.md');
  const context = files.map(f => {
    const name = f.replace('.md', '');
    const content = fs.readFileSync(path.join(WIKI_DIR, f), 'utf-8');
    return `=== ${name} ===\n${content}`;
  }).join('\n\n');

  const msg = await anthropic.messages.create(
    {
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 4096,
      system: [{
        type: 'text',
        text: `당신은 마더스마일 주식회사(유아용품 일본법인) 지식 위키의 AI 어시스턴트입니다.
질문 언어(한국어·일본어·영어 등)에 관계없이 반드시 한국어로 답변하세요.
아래 위키 내용만을 근거로 답변하세요.
위키에 없는 내용은 "위키에 해당 정보가 없습니다"라고 답하세요.
답변은 간결하고 구체적으로, 수치가 있으면 반드시 포함하세요.

위키 내용:
${context}`,
        cache_control: { type: 'ephemeral' }
      }],
      messages: [{ role: 'user', content: `질문: ${question}` }]
    },
    { headers: { 'anthropic-beta': 'prompt-caching-2024-07-31' } }
  );
  return msg.content[0].text;
}

// ── 내보내기 유틸 ──────────────────────────────────
function cleanText(text) {
  return (text || '')
    .replace(/\[\[raw\/[^\]]+\]\]/g, '')
    .replace(/\[\[([^\]|]+)\|([^\]]+)\]\]/g, '$2')
    .replace(/\[\[([^\]]+)\]\]/g, '$1')
    .replace(/[*_`~]/g, '')
    .replace(/\[!\w+\][^\n]*/g, '')
    .replace(/^>\s?/gm, '')
    .trim();
}

function getBlockquoteText(token) {
  if (!token) return '';
  if (typeof token.text === 'string') return token.text;
  if (Array.isArray(token.tokens)) return token.tokens.map(t => getBlockquoteText(t)).join(' ');
  return '';
}

async function generateDocx(pageName, tokens) {
  const children = [];

  const headingMap = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6,
  };

  for (const token of tokens) {
    if (token.type === 'heading') {
      children.push(new Paragraph({ text: cleanText(token.text), heading: headingMap[token.depth] || HeadingLevel.HEADING_3 }));
    } else if (token.type === 'paragraph') {
      const text = cleanText(token.text);
      if (text) children.push(new Paragraph({ text }));
    } else if (token.type === 'blockquote') {
      const raw = getBlockquoteText(token);
      const text = cleanText(raw);
      if (text) children.push(new Paragraph({
        children: [new TextRun({ text, italics: true, color: '555555' })],
        indent: { left: 720 },
        spacing: { before: 100, after: 100 }
      }));
    } else if (token.type === 'list') {
      for (const item of token.items) {
        const text = cleanText(item.text);
        if (text) children.push(new Paragraph({
          text,
          bullet: { level: token.ordered ? 0 : 0 },
        }));
      }
    } else if (token.type === 'table') {
      const rows = [];
      // 헤더
      rows.push(new TableRow({
        tableHeader: true,
        children: token.header.map(h => new TableCell({
          shading: { type: ShadingType.SOLID, color: 'E0E7FF' },
          children: [new Paragraph({ children: [new TextRun({ text: cleanText(h.text), bold: true })] })],
        }))
      }));
      // 데이터
      for (const row of token.rows) {
        rows.push(new TableRow({
          children: row.map(cell => new TableCell({
            children: [new Paragraph({ text: cleanText(cell.text) })],
          }))
        }));
      }
      children.push(new Table({ rows, width: { size: 9000, type: WidthType.DXA } }));
      children.push(new Paragraph({}));
    } else if (token.type === 'code') {
      children.push(new Paragraph({
        children: [new TextRun({ text: token.text, font: 'Courier New', size: 20, color: '333333' })],
        indent: { left: 720 },
        spacing: { before: 60, after: 60 }
      }));
    } else if (token.type === 'space') {
      children.push(new Paragraph({}));
    }
  }

  const doc = new Document({
    sections: [{ properties: {}, children }]
  });
  return Packer.toBuffer(doc);
}

async function generateXlsx(pageName, tokens) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'MSCL 지식 위키';

  // 개요 시트
  const sheet = workbook.addWorksheet('개요');
  sheet.getColumn(1).width = 16;
  sheet.getColumn(2).width = 70;

  const titleRow = sheet.addRow([pageName]);
  titleRow.getCell(1).font = { bold: true, size: 16 };
  sheet.addRow([]);

  for (const token of tokens) {
    if (token.type === 'heading') {
      const text = cleanText(token.text);
      const prefix = '#'.repeat(token.depth) + ' ';
      const r = sheet.addRow([prefix + text]);
      r.getCell(1).font = { bold: true, size: Math.max(10, 15 - token.depth * 2), color: { argb: 'FF1a1a2e' } };
      sheet.mergeCells(`A${r.number}:B${r.number}`);
    } else if (token.type === 'paragraph') {
      const text = cleanText(token.text);
      if (text) {
        const r = sheet.addRow(['', text]);
        r.getCell(2).alignment = { wrapText: true };
      }
    } else if (token.type === 'blockquote') {
      const text = cleanText(getBlockquoteText(token));
      if (text) {
        const r = sheet.addRow(['💡', text]);
        r.getCell(1).font = { bold: true };
        r.getCell(2).alignment = { wrapText: true };
        r.getCell(2).font = { italic: true, color: { argb: 'FF555555' } };
      }
    } else if (token.type === 'list') {
      for (const item of token.items) {
        const text = cleanText(item.text);
        if (text) {
          const r = sheet.addRow(['•', text]);
          r.getCell(2).alignment = { wrapText: true };
        }
      }
    } else if (token.type === 'space') {
      sheet.addRow([]);
    }
  }

  // 표 시트
  let tableNum = 1;
  for (const token of tokens) {
    if (token.type !== 'table') continue;
    const tSheet = workbook.addWorksheet(`표 ${tableNum++}`);
    const headers = token.header.map(h => cleanText(h.text));
    tSheet.columns = headers.map(() => ({ width: 22 }));
    const hRow = tSheet.addRow(headers);
    hRow.eachCell(cell => {
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E7FF' } };
      cell.alignment = { wrapText: true };
    });
    for (const row of token.rows) {
      const dataRow = tSheet.addRow(row.map(c => cleanText(c.text)));
      dataRow.eachCell(cell => { cell.alignment = { wrapText: true }; });
    }
  }

  return workbook.xlsx.writeBuffer();
}

async function generatePptx(pageName, tokens) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.title = pageName;

  const TITLE_COLOR = '1a1a2e';
  const BODY_COLOR = '333333';
  const ACCENT = '4a6fa5';

  // 타이틀 슬라이드
  const titleSlide = pptx.addSlide();
  titleSlide.background = { color: '1a1a2e' };
  titleSlide.addText('📚 MSCL 지식 위키', {
    x: 0.5, y: 1.0, w: 12.33, h: 0.7,
    fontSize: 20, color: '9999cc', align: 'center'
  });
  titleSlide.addText(pageName, {
    x: 0.5, y: 1.9, w: 12.33, h: 1.8,
    fontSize: 40, bold: true, color: 'ffffff', align: 'center', wrap: true
  });

  // takeaway 추출 (첫 blockquote)
  const takeawayToken = tokens.find(t => t.type === 'blockquote');
  if (takeawayToken) {
    const takeaway = cleanText(getBlockquoteText(takeawayToken));
    if (takeaway) {
      titleSlide.addText(takeaway, {
        x: 1, y: 3.8, w: 11.33, h: 2.0,
        fontSize: 16, color: 'aaaadd', align: 'center', italic: true, wrap: true
      });
    }
  }

  // 섹션별 슬라이드
  let currentSlide = null;
  let contentRows = [];

  function flushSlide(slide, rows) {
    if (!slide || rows.length === 0) return;
    // 표가 있으면 별도 처리
    const tableData = rows.filter(r => r.type === 'table');
    const textRows = rows.filter(r => r.type !== 'table');

    if (textRows.length > 0) {
      const textArr = textRows.map(r => ({
        text: r.text + '\n',
        options: { fontSize: r.fontSize || 14, bold: r.bold || false, bullet: r.bullet ? { type: 'bullet' } : false, color: r.color || BODY_COLOR, breakLine: false }
      }));
      slide.addText(textArr, { x: 0.5, y: 1.6, w: 12.33, h: 4.8, valign: 'top', wrap: true });
    }

    if (tableData.length > 0) {
      // 테이블 슬라이드 추가
      for (const td of tableData) {
        const ts = pptx.addSlide();
        ts.addText(td.title || '표', { x: 0.5, y: 0.2, w: 12.33, h: 0.9, fontSize: 22, bold: true, color: TITLE_COLOR });
        const tbl = [];
        tbl.push(td.headers.map(h => ({ text: h, options: { bold: true, fill: 'E0E7FF', color: TITLE_COLOR } })));
        for (const row of td.rows) {
          tbl.push(row.map(c => ({ text: c, options: { color: BODY_COLOR } })));
        }
        ts.addTable(tbl, {
          x: 0.5, y: 1.3, w: 12.33,
          border: { type: 'solid', pt: 0.5, color: 'dddddd' },
          colW: td.headers.map(() => 12.33 / td.headers.length),
          fontSize: 12,
        });
      }
    }
  }

  let sectionTitle = '';
  for (const token of tokens) {
    if (token.type === 'heading' && token.depth <= 2) {
      flushSlide(currentSlide, contentRows);
      contentRows = [];
      sectionTitle = cleanText(token.text);
      currentSlide = pptx.addSlide();
      currentSlide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.12, h: 7.5, fill: { color: ACCENT } });
      currentSlide.addText(sectionTitle, {
        x: 0.5, y: 0.25, w: 12.33, h: 1.1,
        fontSize: 26, bold: true, color: TITLE_COLOR
      });
    } else if (token.type === 'heading' && token.depth >= 3) {
      contentRows.push({ text: cleanText(token.text), bold: true, fontSize: 15, color: ACCENT });
    } else if (token.type === 'paragraph') {
      const text = cleanText(token.text);
      if (text) contentRows.push({ text, fontSize: 13 });
    } else if (token.type === 'blockquote') {
      const text = cleanText(getBlockquoteText(token));
      if (text) contentRows.push({ text: '💡 ' + text, fontSize: 12, color: '555555' });
    } else if (token.type === 'list') {
      for (const item of token.items) {
        const text = cleanText(item.text);
        if (text) contentRows.push({ text, bullet: true, fontSize: 13 });
      }
    } else if (token.type === 'table' && currentSlide) {
      contentRows.push({
        type: 'table',
        title: sectionTitle,
        headers: token.header.map(h => cleanText(h.text)),
        rows: token.rows.map(row => row.map(c => cleanText(c.text)))
      });
    }
  }
  flushSlide(currentSlide, contentRows);

  return pptx.write({ outputType: 'nodebuffer' });
}
// ─────────────────────────────────────────────────

// 라우트
app.get('/', (req, res) => {
  const sidebar = parseSidebar();
  res.render('home', { sidebar });
});

app.post('/api/translate', async (req, res) => {
  const text = (req.body.text || '').trim();
  if (!text || !anthropic) return res.json({ result: null });
  try {
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 4096,
      messages: [{ role: 'user', content: `以下の韓国語テキストを自然な日本語に翻訳してください。マークダウン形式（表・箇条書き・見出しなど）をそのまま保持してください。翻訳結果のみ出力し、説明は不要です。\n\n${text}` }]
    });
    res.json({ result: msg.content[0].text });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/search', (req, res) => {
  const q = (req.query.q || '').trim();
  res.json(q ? search(q) : []);
});

app.get('/api/ask', async (req, res) => {
  const q = (req.query.q || '').trim();
  if (!q) return res.json({ answer: null });
  if (!anthropic) return res.json({ answer: null, noKey: true });
  try {
    const answer = await askAI(q);
    res.json({ answer });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/wiki/:page', (req, res) => {
  const pageName = decodeURIComponent(req.params.page);
  const html = renderPage(pageName);
  const sidebar = parseSidebar();
  if (!html) {
    return res.status(404).render('layout', {
      title: '페이지 없음',
      content: '<p class="not-found">요청한 페이지를 찾을 수 없습니다.</p>',
      sidebar, currentPage: pageName
    });
  }
  res.render('layout', {
    title: pageName === 'index' ? 'MSCL 지식 위키' : pageName,
    content: html, sidebar, currentPage: pageName
  });
});

app.get('/graph', (req, res) => {
  const sidebar = parseSidebar();
  res.render('graph', { sidebar, currentPage: '__graph__' });
});

app.get('/api/graph', (req, res) => {
  res.json(buildGraph());
});

app.get('/search', (req, res) => {
  const q = (req.query.q || '').trim();
  const sidebar = parseSidebar();
  const results = q ? search(q) : [];
  res.render('search', { sidebar, currentPage: '__search__', q, results });
});

app.get('/raw/:filename', (req, res) => {
  const filename = decodeURIComponent(req.params.filename);
  const filePath = path.join(__dirname, 'raw', filename);
  if (!fs.existsSync(filePath)) return res.status(404).send('파일 없음');
  res.download(filePath, filename);
});

app.get('/api/export/:page', async (req, res) => {
  const pageName = decodeURIComponent(req.params.page);
  const format = (req.query.format || 'docx').toLowerCase();
  const filePath = path.join(WIKI_DIR, `${pageName}.md`);
  if (!fs.existsSync(filePath)) return res.status(404).send('페이지 없음');

  const markdown = fs.readFileSync(filePath, 'utf-8');
  const tokens = marked.lexer(markdown);
  const safeName = pageName.replace(/[/\\:*?"<>|]/g, '_');

  try {
    if (format === 'docx') {
      const buf = await generateDocx(pageName, tokens);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(safeName)}.docx`);
      res.send(buf);
    } else if (format === 'xlsx') {
      const buf = await generateXlsx(pageName, tokens);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(safeName)}.xlsx`);
      res.send(buf);
    } else if (format === 'pptx') {
      const buf = await generatePptx(pageName, tokens);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
      res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(safeName)}.pptx`);
      res.send(buf);
    } else {
      res.status(400).send('지원하지 않는 형식입니다 (docx, xlsx, pptx)');
    }
  } catch (e) {
    console.error('Export error:', e);
    res.status(500).send('내보내기 실패: ' + e.message);
  }
});

app.get('/api/status', (req, res) => {
  res.json({
    version: '2026-05-14-v6',
    hasApiKey: !!process.env.ANTHROPIC_API_KEY,
    wikiPages: fs.readdirSync(WIKI_DIR).filter(f => f.endsWith('.md')).length
  });
});

app.listen(PORT, () => console.log(`MSCL 위키 서버 실행 중: http://localhost:${PORT}`));
