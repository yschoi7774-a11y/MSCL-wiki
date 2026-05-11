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
  // [[pagename|display text]]
  html = html.replace(/\[\[([^\]|]+)\|([^\]]+)\]\]/g, (_, page, display) => {
    return `<a href="/wiki/${encodeURIComponent(page)}" class="wikilink">${display}</a>`;
  });
  // [[pagename]]
  html = html.replace(/\[\[([^\]]+)\]\]/g, (_, page) => {
    return `<a href="/wiki/${encodeURIComponent(page)}" class="wikilink">${page}</a>`;
  });
  return html;
}

function processCallouts(markdown) {
  // Collapsible: > [!type]- Title\n> content
  markdown = markdown.replace(
    /^> \[!(\w+)\]-\s*(.*)\n((?:^>.*\n?)*)/gm,
    (_, type, title, body) => {
      const content = body.replace(/^> ?/gm, '').trim();
      return `<details class="callout callout-${type.toLowerCase()}"><summary><span class="callout-type">${type.toUpperCase()}</span>${title ? ' ' + title : ''}</summary><div class="callout-content">\n\n${content}\n\n</div></details>\n`;
    }
  );
  // Regular: > [!type] Title\n> content
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
    if (h2) {
      current = { name: h2[1].trim(), pages: [] };
      categories.push(current);
      continue;
    }
    const link = line.match(/^- \[\[([^\]|]+)(?:\|[^\]]+)?\]\]/);
    if (link && current) {
      const pageName = link[1].trim();
      const displayMatch = line.match(/\[\[[^\]|]+\|([^\]]+)\]\]/);
      const dashMatch = line.match(/\]\] — (.+)/);
      current.pages.push({
        name: pageName,
        display: displayMatch ? displayMatch[1] : pageName,
        desc: dashMatch ? dashMatch[1] : ''
      });
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

app.get('/', (req, res) => res.redirect('/wiki/index'));

app.get('/wiki/:page', (req, res) => {
  const pageName = decodeURIComponent(req.params.page);
  const html = renderPage(pageName);
  const sidebar = parseSidebar();

  if (!html) {
    return res.status(404).render('layout', {
      title: '페이지 없음',
      content: '<p class="not-found">요청한 페이지를 찾을 수 없습니다.</p>',
      sidebar,
      currentPage: pageName
    });
  }

  res.render('layout', {
    title: pageName === 'index' ? 'MSCL 지식 위키' : pageName,
    content: html,
    sidebar,
    currentPage: pageName
  });
});

app.listen(PORT, () => console.log(`MSCL 위키 서버 실행 중: http://localhost:${PORT}`));
