const express = require('express');
const fs = require('fs');
const path = require('path');
const { marked } = require('marked');

const app = express();
const PORT = process.env.PORT || 3000;
const WIKI_DIR = path.join(__dirname, 'wiki');

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

  const files = fs.readdirSync(WIKI_DIR).filter(f => f.endsWith('.md') && f !== 'log.md');
  const nodes = [];
  const links = [];
  const pageSet = new Set(files.map(f => f.replace('.md', '')));

  files.forEach(file => {
    const name = file.replace('.md', '');
    const content = fs.readFileSync(path.join(WIKI_DIR, file), 'utf-8');
    nodes.push({ id: name, category: categoryMap[name] || '기타' });
    extractLinks(content).forEach(target => {
      if (pageSet.has(target) && target !== name) {
        links.push({ source: name, target });
      }
    });
  });

  return { nodes, links };
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

      return { name: page.name, score, snippet };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);
}

// 라우트
app.get('/', (req, res) => res.redirect('/wiki/index'));

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

app.listen(PORT, () => console.log(`MSCL 위키 서버 실행 중: http://localhost:${PORT}`));
