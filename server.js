require('dotenv').config();
const crypto = require('crypto');
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
app.get('/logout', (req, res) => { req.session.destroy(); res.redirect('/login'); })
app.post('/api/chat/clear', (req, res) => { req.session.chatHistory = []; res.json({ ok: true }); });;

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

// AI 질문 답변 (대화 이력 포함)
function buildWikiContext() {
  const files = fs.readdirSync(WIKI_DIR).filter(f => f.endsWith('.md') && f !== 'log.md' && f !== 'index.md');
  return files.map(f => `=== ${f.replace('.md', '')} ===\n${fs.readFileSync(path.join(WIKI_DIR, f), 'utf-8')}`).join('\n\n');
}

async function askAI(messages) {
  if (!anthropic) return null;
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
이전 대화 맥락을 반영해 자연스럽게 이어서 답변하세요.

위키 내용:
${buildWikiContext()}`,
        cache_control: { type: 'ephemeral' }
      }],
      messages
    },
    { headers: { 'anthropic-beta': 'prompt-caching-2024-07-31' } }
  );
  return msg.content[0].text;
}

// ── AI 파일 생성 캐시 (1시간 TTL) ─────────────────
const exportCache = new Map();
setInterval(() => {
  const now = Date.now();
  for (const [id, e] of exportCache) if (e.expires < now) exportCache.delete(id);
}, 3600000);

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
// ── JSON 구조 → PPTX (리디자인) ───────────────────
async function generatePptxFromJSON(s) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; // 13.33 × 7.5 in
  pptx.title = s.title || 'MSCL';

  const C = {
    navy:   '0f1729', navyMid: '1a1a2e', blue:  '4a6fa5',
    blueLt: '6b93c4', gold:    'e8b86d', white: 'ffffff',
    offWht: 'f8fafc', gray:    '64748b', grayLt:'e2e8f0',
    dark:   '1e293b', mid:     '475569',
  };
  const W = 13.33, H = 7.5;

  function addFooter(sl, num) {
    sl.addShape(pptx.ShapeType.rect, { x: 0, y: H - 0.36, w: W, h: 0.36, fill: { color: C.navyMid }, line: { color: C.navyMid } });
    sl.addText('MOTHERSMILE', { x: 0.35, y: H - 0.3, w: 5, h: 0.24, fontSize: 8.5, color: C.blueLt, bold: false });
    if (num) sl.addText(String(num), { x: W - 0.9, y: H - 0.3, w: 0.6, h: 0.24, fontSize: 8.5, color: C.blueLt, align: 'right' });
  }

  // ① 타이틀 슬라이드
  const ts = pptx.addSlide();
  ts.background = { color: C.navy };
  // 왼쪽 세로 바
  ts.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.55, h: H, fill: { color: C.blue }, line: { color: C.blue } });
  // 오른쪽 장식 블록
  ts.addShape(pptx.ShapeType.rect, { x: 11.2, y: 4.8, w: 2.13, h: 2.7, fill: { color: C.navyMid }, line: { color: C.navyMid } });
  ts.addShape(pptx.ShapeType.rect, { x: 11.7, y: 4.4, w: 1.63, h: 3.1, fill: { color: C.blue },    line: { color: C.blue } });
  // 브랜드
  ts.addText('MOTHERSMILE  지식 위키', { x: 0.85, y: 0.5, w: 9, h: 0.45, fontSize: 10.5, color: C.blueLt, charSpacing: 2.5 });
  // 구분선
  ts.addShape(pptx.ShapeType.rect, { x: 0.85, y: 1.08, w: 9.5, h: 0.04, fill: { color: C.blue }, line: { color: C.blue } });
  // 메인 타이틀
  ts.addText(s.title || '', { x: 0.85, y: 1.25, w: 10.5, h: 3.2, fontSize: 42, bold: true, color: C.white, wrap: true, valign: 'middle' });
  // 부제목
  if (s.subtitle) ts.addText(s.subtitle, { x: 0.85, y: 4.6, w: 9.5, h: 0.9, fontSize: 16, color: C.blueLt, italic: true, wrap: true });
  // 날짜
  const today = new Date().toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric' });
  ts.addText(today, { x: 0.85, y: H - 0.75, w: 6, h: 0.35, fontSize: 10, color: C.gray });

  // ② 콘텐츠 슬라이드
  let num = 0;
  for (const slide of (s.slides || [])) {
    num++;
    const type = slide.type || (slide.table ? 'table' : slide.bullets ? 'bullets' : 'section');

    if (type === 'section') {
      // 섹션 구분 슬라이드
      const sl = pptx.addSlide();
      sl.background = { color: C.navyMid };
      sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.55, h: H, fill: { color: C.gold }, line: { color: C.gold } });
      sl.addShape(pptx.ShapeType.rect, { x: 0.55, y: H * 0.5 - 0.03, w: W - 0.55, h: 0.05, fill: { color: C.blue }, line: { color: C.blue } });
      sl.addText(String(num).padStart(2, '0'), { x: 1.1, y: 1.6, w: 3.5, h: 1.4, fontSize: 64, bold: true, color: C.blue });
      sl.addText(slide.heading || '', { x: 1.1, y: 3.1, w: 10.8, h: 2.5, fontSize: 32, bold: true, color: C.white, wrap: true });

    } else if (type === 'table') {
      // 표 슬라이드
      const sl = pptx.addSlide();
      sl.background = { color: C.offWht };
      sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 1.45, fill: { color: C.navyMid }, line: { color: C.navyMid } });
      sl.addText(slide.heading || '', { x: 0.5, y: 0.22, w: 12.3, h: 1.0, fontSize: 22, bold: true, color: C.white });
      if (slide.table) {
        const hdrs = slide.table.headers || [];
        const tbl = [];
        tbl.push(hdrs.map(h => ({ text: h, options: { bold: true, color: C.white, fill: C.blue, align: 'center', valign: 'middle', fontSize: 12 } })));
        (slide.table.rows || []).forEach((row, i) => {
          tbl.push(row.map(c => ({ text: String(c), options: { color: C.dark, fill: i % 2 === 0 ? C.white : 'eef2ff', fontSize: 11.5, valign: 'middle' } })));
        });
        const colW = hdrs.map(() => +(12.33 / hdrs.length).toFixed(2));
        sl.addTable(tbl, { x: 0.5, y: 1.6, w: 12.33, rowH: 0.44, border: { type: 'solid', pt: 0.4, color: C.grayLt }, colW });
      }
      addFooter(sl, num);

    } else {
      // 불릿 슬라이드
      const sl = pptx.addSlide();
      sl.background = { color: C.white };
      // 상단 컬러 바
      sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: 0.07, fill: { color: C.blue }, line: { color: C.blue } });
      // 제목 배경
      sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0.07, w: W, h: 1.32, fill: { color: C.offWht }, line: { color: C.offWht } });
      // 제목 왼쪽 강조 바
      sl.addShape(pptx.ShapeType.rect, { x: 0, y: 0.07, w: 0.18, h: 1.32, fill: { color: C.blue }, line: { color: C.blue } });
      sl.addText(slide.heading || '', { x: 0.42, y: 0.17, w: 12.5, h: 1.1, fontSize: 23, bold: true, color: C.navyMid });
      // 구분선
      sl.addShape(pptx.ShapeType.rect, { x: 0.42, y: 1.42, w: 12.5, h: 0.03, fill: { color: C.grayLt }, line: { color: C.grayLt } });
      // 불릿
      if (slide.bullets && slide.bullets.length) {
        const textArr = slide.bullets.map(b => ({
          text: b,
          options: { bullet: { code: '25B8', color: C.blue }, fontSize: 14.5, color: C.dark, paraSpaceAfter: 8, breakLine: true }
        }));
        sl.addText(textArr, { x: 0.5, y: 1.58, w: 12.33, h: 5.5, valign: 'top', wrap: true, lineSpacingMultiple: 1.35 });
      }
      addFooter(sl, num);
    }
  }
  return pptx.write({ outputType: 'nodebuffer' });
}

// ── JSON 구조 → XLSX (리디자인) ───────────────────
async function generateXlsxFromJSON(s) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'MSCL 지식 위키';
  workbook.created = new Date();

  for (const sheet of (s.sheets || [])) {
    const ws = workbook.addWorksheet(sheet.name || '시트');
    const headers = sheet.headers || [];

    // 1행: 타이틀
    ws.mergeCells(1, 1, 1, headers.length || 1);
    const titleCell = ws.getCell(1, 1);
    titleCell.value = s.title || sheet.name || 'MSCL';
    titleCell.font = { bold: true, size: 16, color: { argb: 'FFFFFFFF' } };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0f1729' } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 32;

    // 2행: 빈 행
    ws.addRow([]);

    // 3행: 헤더
    const hRow = ws.addRow(headers);
    hRow.height = 22;
    hRow.eachCell((cell, col) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1a1a2e' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      cell.border = { bottom: { style: 'medium', color: { argb: 'FF4a6fa5' } } };
    });

    // 데이터 행
    (sheet.rows || []).forEach((row, i) => {
      const dataRow = ws.addRow(row);
      dataRow.height = 18;
      const isAlt = i % 2 === 1;
      dataRow.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isAlt ? 'FFf0f4ff' : 'FFffffff' } };
        cell.alignment = { wrapText: true, vertical: 'middle' };
        cell.border = {
          bottom: { style: 'thin', color: { argb: 'FFe2e8f0' } },
          left:   { style: 'thin', color: { argb: 'FFe2e8f0' } },
          right:  { style: 'thin', color: { argb: 'FFe2e8f0' } },
        };
      });
    });

    // 열 너비 자동 조정 (최소 14, 최대 40)
    ws.columns = headers.map((h, i) => ({
      key: String(i),
      width: Math.min(40, Math.max(14, String(h).length + 6))
    }));

    // 헤더 행 고정
    ws.views = [{ state: 'frozen', ySplit: 3 }];
    // 자동 필터
    if (headers.length) ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3, column: headers.length } };
  }
  return workbook.xlsx.writeBuffer();
}

// ── JSON 구조 → DOCX (리디자인) ──────────────────
async function generateDocxFromJSON(s) {
  const { BorderStyle } = require('docx');
  const children = [];

  // 커버 타이틀
  if (s.title) {
    children.push(new Paragraph({
      children: [new TextRun({ text: s.title, bold: true, size: 56, color: '0f1729' })],
      spacing: { before: 400, after: 200 },
    }));
    children.push(new Paragraph({
      children: [new TextRun({ text: '─'.repeat(40), color: '4a6fa5', size: 16 })],
      spacing: { before: 0, after: 600 },
    }));
  }

  for (const section of (s.sections || [])) {
    if (section.heading) {
      children.push(new Paragraph({
        children: [new TextRun({ text: section.heading, bold: true, size: 30, color: '1a1a2e' })],
        spacing: { before: 480, after: 120 },
        border: { bottom: { color: '4a6fa5', space: 4, style: BorderStyle.SINGLE, size: 6 } }
      }));
    }
    for (const p of (section.paragraphs || [])) {
      if (p) children.push(new Paragraph({
        children: [new TextRun({ text: p, size: 22, color: '1e293b' })],
        spacing: { before: 80, after: 80 }
      }));
    }
    for (const b of (section.bullets || [])) {
      if (b) children.push(new Paragraph({
        children: [new TextRun({ text: b, size: 22, color: '1e293b' })],
        bullet: { level: 0 },
        spacing: { before: 40, after: 40 }
      }));
    }
    if (section.table && section.table.headers) {
      const rows = [];
      rows.push(new TableRow({
        tableHeader: true,
        children: section.table.headers.map(h => new TableCell({
          shading: { type: ShadingType.SOLID, color: '1a1a2e' },
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: 'ffffff', size: 20 })] })]
        }))
      }));
      (section.table.rows || []).forEach((row, i) => {
        rows.push(new TableRow({
          children: row.map(c => new TableCell({
            shading: { type: ShadingType.SOLID, color: i % 2 === 0 ? 'ffffff' : 'eef2ff' },
            margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: [new Paragraph({ children: [new TextRun({ text: String(c), size: 20, color: '1e293b' })] })]
          }))
        }));
      });
      children.push(new Table({
        rows,
        width: { size: 9000, type: WidthType.DXA },
      }));
      children.push(new Paragraph({ spacing: { before: 160 } }));
    }
  }

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Malgun Gothic', size: 22 } } }
    },
    sections: [{
      properties: { page: { margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
      children
    }]
  });
  return Packer.toBuffer(doc);
}

// ── AI 파일 생성 ──────────────────────────────────
async function aiGenerateFile(question, wikiContext, format) {
  if (!anthropic) return null;

  const fmtSchemas = {
    pptx: `슬라이드 타입: "section"(챕터 구분), "bullets"(목록), "table"(표)
JSON만 반환:
{"title":"발표 제목","subtitle":"부제목(선택)","slides":[
  {"type":"section","heading":"챕터명"},
  {"type":"bullets","heading":"슬라이드 제목","bullets":["핵심 내용을 문장으로","수치 포함 구체적으로","항목당 1-2줄 권장"]},
  {"type":"table","heading":"슬라이드 제목","table":{"headers":["컬럼1","컬럼2","컬럼3"],"rows":[["값1","값2","값3"]]}}
]}`,
    xlsx: `JSON만 반환:
{"title":"문서 제목","sheets":[{"name":"시트명","headers":["컬럼1","컬럼2","컬럼3"],"rows":[["값1","값2","값3"]]}]}`,
    docx: `JSON만 반환:
{"title":"문서 제목","sections":[
  {"heading":"섹션 제목","paragraphs":["단락 내용"],"bullets":["항목1","항목2"]},
  {"heading":"표 섹션","table":{"headers":["컬럼1","컬럼2"],"rows":[["값1","값2"]]}}
]}`
  };

  const msg = await anthropic.messages.create(
    {
      model: 'claude-sonnet-4-6',
      max_tokens: 4096,
      system: [{
        type: 'text',
        text: `마더스마일 지식 위키 기반으로 사용자 요청에 맞는 파일 구조를 JSON으로만 반환하세요.
JSON 외 어떤 텍스트도 출력하지 마세요. 내용은 구체적이고 실무에 바로 쓸 수 있게 작성하세요.

${fmtSchemas[format]}

위키 내용:
${wikiContext}`,
        cache_control: { type: 'ephemeral' }
      }],
      messages: [{ role: 'user', content: question }]
    },
    { headers: { 'anthropic-beta': 'prompt-caching-2024-07-31' } }
  );

  const text = msg.content[0].text.trim();
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) return null;
  let structure;
  try {
    structure = JSON.parse(jsonMatch[0]);
  } catch (e) {
    // JSON 일부 정제 후 재시도
    const cleaned = jsonMatch[0]
      .replace(/[\u0000-\u001F\u007F]/g, ' ')  // 제어문자 제거
      .replace(/,\s*([}\]])/g, '$1');           // trailing comma 제거
    try { structure = JSON.parse(cleaned); } catch { return null; }
  }

  const fmtMeta = {
    pptx: { gen: generatePptxFromJSON, ct: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', ext: 'pptx' },
    xlsx: { gen: generateXlsxFromJSON, ct: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', ext: 'xlsx' },
    docx: { gen: generateDocxFromJSON, ct: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', ext: 'docx' }
  };
  const { gen, ct, ext } = fmtMeta[format];
  const buffer = await gen(structure);
  const filename = `${(structure.title || 'MSCL문서').replace(/[/\\:*?"<>|]/g, '_')}.${ext}`;
  return { buffer, filename, ct, format: ext, title: structure.title };
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
    // 파일 생성 요청 감지
    const MAKE_RE = /만들|작성|생성|뽑아|정리|export/i;
    const FORM_RE = /양식|서식|템플릿|template|대장|신청서|보고서|체크리스트|명세서|계획서|지침서|매뉴얼/i;
    const PPT_RE  = /ppt|pptx|파워포인트|슬라이드|발표|프레젠테이션/i;
    const XLSX_RE = /엑셀|excel|xlsx|스프레드시트/i;
    const DOCX_RE = /워드|word|docx/i;

    let fileFormat = null;
    if (MAKE_RE.test(q) || FORM_RE.test(q)) {
      if (PPT_RE.test(q))       fileFormat = 'pptx';
      else if (XLSX_RE.test(q)) fileFormat = 'xlsx';
      else if (DOCX_RE.test(q)) fileFormat = 'docx';
      else if (FORM_RE.test(q)) fileFormat = 'xlsx'; // 양식/서식은 기본 Excel
    }

    if (fileFormat) {
      return res.json({ type: 'download', q, format: fileFormat });
    }

    // 세션 대화 이력 관리 (최대 10턴 = 20 메시지)
    if (!req.session.chatHistory) req.session.chatHistory = [];
    req.session.chatHistory.push({ role: 'user', content: q });
    if (req.session.chatHistory.length > 20) req.session.chatHistory = req.session.chatHistory.slice(-20);

    const answer = await askAI(req.session.chatHistory);
    req.session.chatHistory.push({ role: 'assistant', content: answer });

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

app.get('/api/ai-download', async (req, res) => {
  const q = (req.query.q || '').trim();
  const format = (req.query.format || 'pptx').toLowerCase();
  if (!q || !anthropic) return res.status(400).send('잘못된 요청');
  try {
    const files = fs.readdirSync(WIKI_DIR).filter(f => f.endsWith('.md') && f !== 'log.md' && f !== 'index.md');
    const context = files.map(f => `=== ${f.replace('.md','')} ===\n${fs.readFileSync(path.join(WIKI_DIR, f), 'utf-8')}`).join('\n\n');
    const result = await aiGenerateFile(q, context, format);
    if (!result) return res.status(500).send('파일 생성 실패');
    res.setHeader('Content-Type', result.ct);
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(result.filename)}`);
    res.send(result.buffer);
  } catch (e) {
    console.error('ai-download error:', e);
    res.status(500).send('오류: ' + e.message);
  }
});

app.get('/api/temp/:id', (req, res) => {
  const entry = exportCache.get(req.params.id);
  if (!entry || entry.expires < Date.now()) return res.status(404).send('파일이 만료되었습니다 (1시간 유효)');
  res.setHeader('Content-Type', entry.ct);
  res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(entry.filename)}`);
  res.send(entry.buffer);
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
    version: '2026-05-14-v10',
    hasApiKey: !!process.env.ANTHROPIC_API_KEY,
    wikiPages: fs.readdirSync(WIKI_DIR).filter(f => f.endsWith('.md')).length
  });
});

app.listen(PORT, () => console.log(`MSCL 위키 서버 실행 중: http://localhost:${PORT}`));
