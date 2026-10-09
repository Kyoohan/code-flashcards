/* ============================================================
   nbview.js — .ipynb 뷰어와 빈칸

   노트북의 마크다운 · 코드 · 출력을 노트북처럼 보여주고,
   텍스트와 코드 어디에든 빈칸(소스 오프셋 범위)을 둘 수 있게 한다.

   - 이 파일은 index.html 의 메인 스크립트보다 먼저 로드된다.
     index.html 의 전역(sanitizeBlanks, autoBlanks, genId)은 호출 시점에만 참조한다.
   - 노트북은 외부에서 온 신뢰할 수 없는 파일이다. 사이트가 Gist 토큰을 localStorage 에
     들고 있으므로, 출력 HTML 은 허용 목록 방식으로 정화한 뒤에만 DOM 에 넣는다.
   ============================================================ */
'use strict';

const NBV_MAX_TEXT = 8000;      // 출력 텍스트 한 덩어리 상한(글자)
const NBV_MAX_HTML = 40000;     // html 출력 상한 — 넘으면 text/plain 으로 대체
const NBV_MAX_IMG  = 250000;    // 이미지 base64 길이 상한 (약 190KB)

const nbvEsc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/* ────────────────────────────────────────────────────────────
   1. 출력 · 셀 파싱 (nbformat 4)
   ──────────────────────────────────────────────────────────── */
const nbvJoin = v => Array.isArray(v) ? v.join('') : String(v == null ? '' : v);
const nbvStripAnsi = s => String(s).replace(/\u001b\[[0-9;?]*[ -\/]*[@-~]/g, '');
function nbvTrunc(s, max) { s = String(s); return s.length > max ? s.slice(0, max) + '\n… (이하 생략)' : s; }
/** tqdm 같은 진행 막대는 \r 로 줄을 덮어쓴다 — 마지막 상태만 남긴다 */
function nbvCollapseCR(s) {
  return String(s).split('\n').map(l => l.includes('\r') ? (l.split('\r').filter(Boolean).pop() || '') : l).join('\n');
}

const NBV_IMG_MIME = { 'image/png': 1, 'image/jpeg': 1, 'image/gif': 1, 'image/webp': 1 };
function nbvImageOut(mime, raw) {
  const data = nbvJoin(raw).replace(/\s+/g, '');
  if (!/^[A-Za-z0-9+/=]+$/.test(data)) return null;
  return { kind: 'image', mime, data, size: data.length };
}

/**
 * 마크다운 셀의 첨부 이미지 (nbformat 의 cell.attachments, .docx 의 그림).
 * 저장 형태: cell.att = { 이름: { mime, data(base64) } }. 본문에서는 ![설명](attachment:이름) 으로 가리킨다.
 * refs 를 주면 본문이 실제로 가리키는 것만 남긴다.
 */
function nbvCleanAtt(att, refs) {
  if (!att || typeof att !== 'object') return null;
  const out = {};
  let n = 0;
  Object.keys(att).forEach(name => {
    if (refs && !refs.has(name)) return;
    const f = att[name];
    if (!f || !NBV_IMG_MIME[f.mime] || typeof f.data !== 'string' || !/^[A-Za-z0-9+/=]+$/.test(f.data) || f.data.length > NBV_MAX_IMG) return;
    out[name] = { mime: f.mime, data: f.data };
    n++;
  });
  return n ? out : null;
}
const nbvAttRefs = src => new Set([...String(src).matchAll(/!\[[^\]]*\]\(attachment:([^)\s]+)\)/g)].map(m => m[1]));
/** nbformat attachments({이름: {mime: base64}}) → {이름: {mime, data}} */
function nbvFromNbAttachments(a) {
  if (!a || typeof a !== 'object') return null;
  const out = {};
  Object.keys(a).forEach(name => {
    const bundle = a[name] || {};
    const mime = Object.keys(bundle).find(m => NBV_IMG_MIME[m]);
    if (mime) out[name] = { mime, data: nbvJoin(bundle[mime]).replace(/\s+/g, '') };
  });
  return out;
}

function nbvPickMime(d) {
  for (const mime of ['image/png', 'image/jpeg', 'image/gif']) {
    if (d[mime]) { const o = nbvImageOut(mime, d[mime]); if (o) return o; }
  }
  if (d['image/svg+xml']) {
    const text = nbvJoin(d['image/svg+xml']);
    return { kind: 'svg', text, size: text.length };
  }
  const plain = d['text/plain'] != null ? nbvTrunc(nbvJoin(d['text/plain']), NBV_MAX_TEXT) : '';
  if (d['text/html']) {
    const html = nbvJoin(d['text/html']);
    if (html.length <= NBV_MAX_HTML) return { kind: 'html', html };
    if (plain) return { kind: 'text', name: 'stdout', text: plain };      // 너무 큰 표는 텍스트로
    return null;
  }
  if (d['text/markdown']) return { kind: 'md', text: nbvTrunc(nbvJoin(d['text/markdown']), NBV_MAX_TEXT) };
  if (plain.trim()) return { kind: 'text', name: 'stdout', text: plain };
  return null;
}

function nbvParseOutputs(outs) {
  const res = [];
  (Array.isArray(outs) ? outs : []).forEach(o => {
    if (!o || typeof o !== 'object') return;
    if (o.output_type === 'stream') {
      const t = nbvTrunc(nbvCollapseCR(nbvJoin(o.text)), NBV_MAX_TEXT);
      if (t.trim()) res.push({ kind: 'text', name: o.name === 'stderr' ? 'stderr' : 'stdout', text: t });
    } else if (o.output_type === 'error') {
      const tb = (Array.isArray(o.traceback) ? o.traceback : []).map(nbvStripAnsi).join('\n');
      res.push({ kind: 'error', text: nbvTrunc(tb || `${o.ename || 'Error'}: ${o.evalue || ''}`, NBV_MAX_TEXT) });
    } else if (o.output_type === 'execute_result' || o.output_type === 'display_data') {
      const item = nbvPickMime(o.data && typeof o.data === 'object' ? o.data : {});
      if (item) { if (o.output_type === 'execute_result') item.result = true; res.push(item); }
    }
  });
  return res;
}

/** 마크다운 셀을 제목 줄 기준으로 쪼갠다 (코드 펜스 안의 # 은 제목이 아니다) */
function nbvSplitMd(source) {
  const chunks = [];
  let cur = [], fence = null;
  source.split('\n').forEach(l => {
    const f = l.match(/^\s{0,3}(```|~~~)/);
    if (f) { if (!fence) fence = f[1]; else if (f[1] === fence) fence = null; }
    const isHead = !fence && !f && /^\s{0,3}#{1,6}\s+\S/.test(l);
    if (isHead && cur.some(x => x.trim())) { chunks.push(cur.join('\n')); cur = []; }
    cur.push(l);
  });
  if (cur.some(x => x.trim())) chunks.push(cur.join('\n'));
  return chunks.map(c => c.replace(/^\n+/, '').replace(/\s+$/, ''));
}

/** .ipynb(JSON) → 정규화된 셀 목록. 마크다운은 제목 단위로 미리 쪼개 둔다 */
function nbvParseCells(nb) {
  const meta = nb && nb.metadata || {};
  const lang = String((meta.language_info && meta.language_info.name) || (meta.kernelspec && meta.kernelspec.language) || 'python').toLowerCase();
  const cells = [];
  (Array.isArray(nb && nb.cells) ? nb.cells : []).forEach(c => {
    if (!c || typeof c !== 'object') return;
    const source = nbvJoin(c.source).replace(/\r\n?/g, '\n').replace(/\s+$/, '');
    if (c.cell_type === 'markdown') {
      const att = nbvFromNbAttachments(c.attachments);
      nbvSplitMd(source).forEach(chunk => {
        if (!chunk.trim()) return;
        const cell = { type: 'markdown', source: chunk, blanks: [] };
        const a = att && nbvCleanAtt(att, nbvAttRefs(chunk));     // 쪼갠 조각마다 자기가 쓰는 그림만
        if (a) cell.att = a;
        cells.push(cell);
      });
    } else if (c.cell_type === 'code') {
      const outputs = nbvParseOutputs(c.outputs);
      if (!source.trim() && !outputs.length) return;
      cells.push({ type: 'code', source, blanks: [], exec: c.execution_count == null ? null : c.execution_count, outputs });
    }
  });
  return { lang, cells };
}

/** 저장된 카드의 셀을 검증·정규화 (손상된 데이터나 외부 JSON 방어) */
function nbvValidOutput(o) {
  if (!o || typeof o !== 'object') return false;
  if (o.kind === 'text' || o.kind === 'error' || o.kind === 'md') return typeof o.text === 'string';
  if (o.kind === 'html') return typeof o.html === 'string';
  if (o.kind === 'svg') return typeof o.text === 'string';
  if (o.kind === 'image') return !!NBV_IMG_MIME[o.mime] && typeof o.data === 'string' && /^[A-Za-z0-9+/=]+$/.test(o.data);
  return false;
}
/* ── 출력의 빈칸 ──────────────────────────────────────────
   텍스트·오류·마크다운 출력은 o.text, HTML 출력(pandas 표 등)은 정화된 HTML 의 '글자만 이어붙인 텍스트'를
   기준으로 오프셋을 잡는다. 이미지·SVG 는 글자가 없어 빈칸을 둘 수 없다. */
const nbvHtmlTextMemo = new Map();
function nbvHtmlText(html) {
  let t = nbvHtmlTextMemo.get(html);
  if (t === undefined) {
    t = nbvSanitize(html).textContent;
    if (nbvHtmlTextMemo.size > 200) nbvHtmlTextMemo.clear();
    nbvHtmlTextMemo.set(html, t);
  }
  return t;
}
function nbvOutText(o) {
  if (!o) return null;
  if (o.kind === 'text' || o.kind === 'error' || o.kind === 'md') return o.text;
  if (o.kind === 'html') return nbvHtmlText(o.html);
  return null;
}
const nbvOutBlankable = o => nbvOutText(o) != null;
const nbvOutEditable = o => !!o && (o.kind === 'text' || o.kind === 'error' || o.kind === 'md');
/** 셀 하나가 가진 빈칸 수 (소스 + 출력) */
function nbvCellBlankCount(cell) {
  return (cell.blanks || []).length + (cell.outputs || []).reduce((a, o) => a + (o.blanks || []).length, 0);
}

function nbvNormalizeOutput(o) {
  if (!Array.isArray(o.blanks) || !o.blanks.length) {
    if (!('blanks' in o)) return o;
    const { blanks, ...rest } = o; return rest;
  }
  const text = nbvOutText(o);
  const bl = text == null ? [] : sanitizeBlanks(o.blanks, text);
  if (bl.length) return { ...o, blanks: bl };
  const { blanks, ...rest } = o;
  return rest;
}
function nbvNormalizeCell(c) {
  if (!c || typeof c !== 'object') return null;
  const type = c.type === 'code' ? 'code' : 'markdown';
  const source = String(c.source == null ? '' : c.source);
  const cell = { type, source, blanks: sanitizeBlanks(c.blanks, source) };
  if (type === 'markdown') { const a = nbvCleanAtt(c.att); if (a) cell.att = a; }
  if (type === 'code') {
    cell.exec = c.exec == null ? null : c.exec;
    cell.outputs = (Array.isArray(c.outputs) ? c.outputs : []).filter(nbvValidOutput).map(nbvNormalizeOutput);
  }
  return cell;
}

/* ────────────────────────────────────────────────────────────
   2. 섹션 분리 → 카드
   ──────────────────────────────────────────────────────────── */
function nbvPlain(s) {
  return String(s).replace(/!?\[([^\]]*)\]\([^)]*\)/g, '$1').replace(/[*_`~]/g, '').replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
}
function nbvHeadingOf(src) {
  const first = String(src).split('\n').find(l => l.trim());
  const m = first && first.match(/^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$/);
  return m ? { level: m[1].length, text: nbvPlain(m[2]) } : null;
}

/** mode: 'section'(제목 단위) | 'cell'(코드 셀 단위) | 'whole'(노트북 전체) */
function nbvGroup(cells, mode, depth, fallback) {
  const groups = [];
  if (mode === 'whole') {
    return cells.length ? [{ title: fallback, cells: cells.slice() }] : [];
  }
  if (mode === 'cell') {
    let heading = '', pend = [], n = 0;
    cells.forEach(c => {
      pend.push(c);
      if (c.type === 'markdown') { const h = nbvHeadingOf(c.source); if (h) heading = h.text; }
      else { n++; groups.push({ title: `${heading || fallback} · 셀 ${n}`, cells: pend }); pend = []; }
    });
    if (pend.length) groups.push({ title: heading || fallback, cells: pend });
    return groups;
  }
  let cur = null;
  cells.forEach(c => {
    const h = c.type === 'markdown' ? nbvHeadingOf(c.source) : null;
    if (!cur || (h && h.level <= depth)) { cur = { title: h ? h.text : fallback, cells: [] }; groups.push(cur); }
    cur.cells.push(c);
  });
  return groups;
}

/** opts: { blank: 'off'|'low'|'mid'|'high', outputs: 'all'|'noimg'|'none' } */
function nbvMakeCard(group, lang, opts) {
  const cells = group.cells.map(c => {
    const out = { type: c.type, source: c.source, blanks: [] };
    if (c.att && opts.outputs === 'all') out.att = c.att;      // 그림 빼기를 고르면 문서 그림도 뺀다
    if (c.type === 'code') {
      out.exec = c.exec == null ? null : c.exec;
      out.outputs = [];
      if (opts.outputs !== 'none') {
        (c.outputs || []).forEach(o => {
          if (o.kind === 'image' || o.kind === 'svg') {
            if (opts.outputs === 'noimg') return;
            if ((o.size || 0) > NBV_MAX_IMG) { out.outputs.push({ kind: 'text', name: 'stdout', text: '[이미지 생략 — 용량 초과]' }); return; }
          }
          out.outputs.push(o);
        });
      }
      if (opts.blank && opts.blank !== 'off') out.blanks = autoBlanks(c.source, opts.blank);
    }
    return out;
  });
  return { id: genId('card'), kind: 'notebook', definition: String(group.title || '노트북').slice(0, 200), term: '', blanks: [], lang, cells };
}

function nbvCardStats(card) {
  let md = 0, code = 0, outs = 0, blanks = 0;
  (card.cells || []).forEach(c => {
    if (c.type === 'code') { code++; outs += (c.outputs || []).length; } else md++;
    blanks += nbvCellBlankCount(c);
  });
  return { md, code, outs, blanks, bytes: JSON.stringify(card.cells || []).length };
}

/* ────────────────────────────────────────────────────────────
   3. 코드 하이라이트 · 마크다운 렌더링
   ──────────────────────────────────────────────────────────── */
function nbvHighlight(src, lang) {
  if (!window.hljs) return nbvEsc(src);
  try {
    if (lang && hljs.getLanguage(lang)) return hljs.highlight(src, { language: lang, ignoreIllegals: true }).value;
    return hljs.highlightAuto(src).value;
  } catch (e) { return nbvEsc(src); }
}

let nbvAtt = null;      // 지금 그리는 마크다운 셀의 첨부 그림
/** 인라인 마크다운. 입력은 원문(이스케이프 전). 먼저 전부 이스케이프하므로 원문 HTML 은 통과하지 못한다 */
function nbvInline(raw) {
  let t = nbvEsc(raw);
  const stash = [];
  const keep = h => { stash.push(h); return '\u0001' + (stash.length - 1) + '\u0002'; };
  t = t.replace(/&lt;br\s*\/?&gt;/gi, () => keep('<br>'));
  t = t.replace(/(`+)([^`\n]+?)\1/g, (_, __, c) => keep('<code class="md-code">' + c + '</code>'));
  t = t.replace(/!\[([^\]]*)\]\(([^)]*)\)/g, (_, alt, url) => {
    // 셀에 첨부된 그림만 그린다 (외부 주소는 불러오지 않는다). 데이터는 nbvCleanAtt 로 검증된 base64
    const m = /^attachment:(.+)$/.exec(url.replace(/&amp;/g, '&'));
    const f = m && nbvAtt && Object.prototype.hasOwnProperty.call(nbvAtt, m[1]) ? nbvAtt[m[1]] : null;
    if (f && NBV_IMG_MIME[f.mime] && /^[A-Za-z0-9+/=]+$/.test(f.data)) return keep(`<img class="md-img" alt="${alt}" loading="lazy" src="data:${f.mime};base64,${f.data}">`);
    return alt ? keep('<span class="md-imgalt">[' + alt + ']</span>') : '';
  });
  t = t.replace(/\[([^\]]+)\]\(([^)\s]+)(?:\s+&quot;[^&]*?&quot;)?\)/g, (m, text, url) => {
    const u = url.replace(/&amp;/g, '&');
    return /^(https?:\/\/|mailto:|#)/i.test(u)
      ? keep('<a href="' + nbvEsc(u) + '" target="_blank" rel="noopener noreferrer">') + text + keep('</a>')
      : text;
  });
  t = t.replace(/\*\*(?=\S)([^*]*?\S)\*\*/g, '<strong>$1</strong>');
  t = t.replace(/(^|[^\w])__(?=\S)([^_]*?\S)__(?!\w)/g, '$1<strong>$2</strong>');
  t = t.replace(/\*(?=\S)([^*\n]*?\S)\*/g, '<em>$1</em>');
  t = t.replace(/(^|[^\w])_(?=\S)([^_\n]*?\S)_(?!\w)/g, '$1<em>$2</em>');
  t = t.replace(/~~(?=\S)([^~]*?\S)~~/g, '<del>$1</del>');
  return t.replace(/\u0001(\d+)\u0002/g, (_, i) => stash[+i]);
}

function nbvMd(src, lang) { return nbvMdBlocks(String(src).replace(/\r\n?/g, '\n').split('\n'), lang); }

function nbvMdBlocks(lines, lang) {
  let out = '', i = 0;
  const blank = l => /^\s*$/.test(l);
  const fenceM = l => l.match(/^\s{0,3}(```|~~~)\s*([\w+#.-]*)/);
  const head = l => /^\s{0,3}#{1,6}\s+\S/.test(l);
  const hr = l => /^\s{0,3}([-*_])(\s*\1){2,}\s*$/.test(l);
  const quote = l => /^\s{0,3}>/.test(l);
  const listM = l => l.match(/^(\s*)([-*+]|\d{1,9}[.)])\s+(.*)$/);
  const tsep = l => /^\s*\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)+\|?\s*$/.test(l) || /^\s*\|\s*:?-{1,}:?\s*\|\s*$/.test(l);
  const tableAt = k => k + 1 < lines.length && lines[k].includes('|') && tsep(lines[k + 1]);
  const cells = l => l.trim().replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim());

  while (i < lines.length) {
    const line = lines[i];
    if (blank(line)) { i++; continue; }
    let m;
    if ((m = fenceM(line))) {
      const fence = m[1], l = (m[2] || '').toLowerCase();
      const buf = [];
      i++;
      while (i < lines.length && !new RegExp('^\\s{0,3}' + fence).test(lines[i])) { buf.push(lines[i]); i++; }
      i++;
      out += '<pre class="md-pre"><code class="hljs">' + nbvHighlight(buf.join('\n'), l || lang) + '</code></pre>';
      continue;
    }
    if ((m = line.match(/^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$/))) {
      const n = m[1].length;
      out += `<h${n} class="md-h md-h${n}">${nbvInline(m[2])}</h${n}>`;
      i++; continue;
    }
    if (hr(line)) { out += '<hr class="md-hr">'; i++; continue; }
    if (quote(line)) {
      const buf = [];
      while (i < lines.length && quote(lines[i])) { buf.push(lines[i].replace(/^\s{0,3}>\s?/, '')); i++; }
      out += '<blockquote class="md-quote">' + nbvMdBlocks(buf, lang) + '</blockquote>';
      continue;
    }
    if (tableAt(i)) {
      const headCells = cells(lines[i]);
      const al = cells(lines[i + 1]).map(s => /^:-+:$/.test(s) ? 'center' : /-+:$/.test(s) ? 'right' : '');
      i += 2;
      const rows = [];
      while (i < lines.length && !blank(lines[i]) && lines[i].includes('|')) { rows.push(cells(lines[i])); i++; }
      const td = (tag, txt, k) => `<${tag}${al[k] ? ` style="text-align:${al[k]}"` : ''}>${nbvInline(txt)}</${tag}>`;
      out += '<div class="md-table-wrap"><table class="md-table"><thead><tr>' + headCells.map((h, k) => td('th', h, k)).join('')
        + '</tr></thead><tbody>' + rows.map(r => '<tr>' + headCells.map((_, k) => td('td', r[k] || '', k)).join('') + '</tr>').join('') + '</tbody></table></div>';
      continue;
    }
    if ((m = listM(line))) {
      const ordered = /\d/.test(m[2]);
      const base = m[1].replace(/\t/g, '    ').length;
      const items = [];
      while (i < lines.length) {
        const lm = listM(lines[i]);
        const ind = lm ? lm[1].replace(/\t/g, '    ').length : 0;
        if (lm && ind < base + 2 && /\d/.test(lm[2]) === ordered) { items.push({ first: lm[3], rest: [] }); i++; continue; }
        if (!items.length) break;
        if (blank(lines[i])) {
          let k = i + 1;
          while (k < lines.length && blank(lines[k])) k++;
          const nl = k < lines.length ? listM(lines[k]) : null;
          const nextInd = k < lines.length ? lines[k].match(/^\s*/)[0].replace(/\t/g, '    ').length : 0;
          if (k < lines.length && ((nl && nl[1].replace(/\t/g, '    ').length >= base && /\d/.test(nl[2]) === ordered) || nextInd > base)) { i = k; continue; }
          break;
        }
        const cur = lines[i].match(/^\s*/)[0].replace(/\t/g, '    ').length;
        if (cur > base || (!lm && cur > 0)) { items[items.length - 1].rest.push(lines[i]); i++; continue; }
        break;
      }
      const tag = ordered ? 'ol' : 'ul';
      out += `<${tag} class="md-list">` + items.map(it => {
        const nonBlank = it.rest.filter(x => x.trim());
        const strip = nonBlank.length ? Math.min(...nonBlank.map(x => x.match(/^\s*/)[0].length)) : 0;
        const body = [it.first, ...it.rest.map(x => x.slice(Math.min(strip, x.match(/^\s*/)[0].length)))];
        return '<li>' + nbvMdBlocks(body, lang) + '</li>';
      }).join('') + `</${tag}>`;
      continue;
    }
    const buf = [line];
    i++;
    while (i < lines.length && !blank(lines[i]) && !fenceM(lines[i]) && !head(lines[i]) && !hr(lines[i]) && !quote(lines[i]) && !listM(lines[i]) && !tableAt(i)) { buf.push(lines[i]); i++; }
    out += '<p class="md-p">' + nbvInline(buf.join('\n')).replace(/ {2,}\n/g, '<br>').replace(/\n/g, ' ') + '</p>';
  }
  return out;
}

/* ────────────────────────────────────────────────────────────
   4. 출력 HTML 정화 (허용 목록)
   ──────────────────────────────────────────────────────────── */
const NBV_ALLOWED = new Set(['TABLE', 'THEAD', 'TBODY', 'TFOOT', 'TR', 'TH', 'TD', 'CAPTION', 'COLGROUP', 'COL', 'DIV', 'SPAN', 'P', 'BR', 'B', 'I', 'U', 'S',
  'STRONG', 'EM', 'CODE', 'PRE', 'UL', 'OL', 'LI', 'DL', 'DT', 'DD', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'SMALL', 'SUB', 'SUP', 'HR', 'BLOCKQUOTE', 'IMG']);
const NBV_DROP = new Set(['SCRIPT', 'STYLE', 'IFRAME', 'OBJECT', 'EMBED', 'FORM', 'INPUT', 'BUTTON', 'SELECT', 'TEXTAREA', 'LINK', 'META', 'TEMPLATE',
  'NOSCRIPT', 'SVG', 'MATH', 'VIDEO', 'AUDIO', 'CANVAS', 'BASE', 'APPLET', 'FRAME', 'FRAMESET', 'HEAD', 'TITLE']);
const NBV_DATA_IMG = /^data:image\/(png|jpe?g|gif|webp);base64,[A-Za-z0-9+/=]+$/;

/** 신뢰할 수 없는 HTML → 안전한 DocumentFragment.
 *  DOMParser 문서는 스크립트가 실행되지 않고 이미지도 불러오지 않는다.
 *  결과는 원본 노드를 옮기는 게 아니라 허용된 태그를 새로 만들어 담으므로, 허용하지 않은 속성(on*, style, href …)은 애초에 존재하지 않는다. */
function nbvSanitize(html) {
  const doc = new DOMParser().parseFromString('<!DOCTYPE html><body>' + html, 'text/html');
  const frag = document.createDocumentFragment();
  (function walk(src, dst) {
    src.childNodes.forEach(n => {
      if (n.nodeType === 3) { dst.appendChild(document.createTextNode(n.nodeValue)); return; }
      if (n.nodeType !== 1) return;
      const tag = n.tagName.toUpperCase();
      if (NBV_DROP.has(tag)) return;
      if (!NBV_ALLOWED.has(tag)) { walk(n, dst); return; }          // 모르는 태그는 벗기고 내용만 둔다
      const el = document.createElement(tag.toLowerCase());
      if (tag === 'IMG') {
        const s = n.getAttribute('src') || '';
        if (!NBV_DATA_IMG.test(s) || s.length > NBV_MAX_IMG) return;  // 외부 주소 이미지는 불러오지 않는다(추적 방지)
        el.setAttribute('src', s);
        const alt = n.getAttribute('alt'); if (alt) el.setAttribute('alt', alt.slice(0, 120));
      }
      ['colspan', 'rowspan', 'align'].forEach(a => {
        if (n.hasAttribute(a)) { const v = n.getAttribute(a); if (/^[\w%-]{1,12}$/.test(v)) el.setAttribute(a, v); }
      });
      walk(n, el);
      dst.appendChild(el);
    });
  })(doc.body, frag);
  return frag;
}

/* ── print(...) 탐지 ────────────────────────────────────────
   편집기에서 print 문마다 '안쪽 내용을 빈칸으로' 버튼을 띄우기 위한 것. */
const NBV_SLASH_LANGS = new Set(['javascript', 'typescript', 'java', 'c', 'cpp', 'c++', 'csharp', 'c#', 'go', 'rust', 'kotlin', 'php', 'swift', 'scala', 'dart']);
/** 문자열·주석 안을 1 로 표시한 마스크. 파이썬 계열에서 // 는 주석이 아니라 나눗셈이므로 언어별로 다르게 본다 */
function nbvMask(src, lang) {
  if (NBV_SLASH_LANGS.has(String(lang || '').toLowerCase()) && typeof codeMask === 'function') return codeMask(src);
  const mask = new Uint8Array(src.length);
  let i = 0;
  while (i < src.length) {
    const c = src[i];
    if (c === '#') { while (i < src.length && src[i] !== '\n') mask[i++] = 1; continue; }
    if (c === '"' || c === "'") {
      const triple = src[i + 1] === c && src[i + 2] === c;
      const qlen = triple ? 3 : 1;
      let j = i + qlen;
      while (j < src.length) {
        if (src[j] === '\\') { j += 2; continue; }
        if (triple ? (src[j] === c && src[j + 1] === c && src[j + 2] === c) : src[j] === c) { j += qlen; break; }
        if (!triple && src[j] === '\n') break;
        j++;
      }
      for (let k = i; k < Math.min(j, src.length); k++) mask[k] = 1;
      i = Math.max(j, i + 1);
      continue;
    }
    i++;
  }
  return mask;
}

/** 소스 안의 print(...) 호출. start/end 는 괄호 안쪽 내용(앞뒤 공백 제외), close 는 ')' 바로 다음 위치.
 *  문자열·주석 안의 print 는 제외하고, print(print(x)) 처럼 겹친 경우는 바깥 것만 찾는다. */
function nbvFindPrints(src, lang) {
  const mask = nbvMask(src, lang);
  const out = [];
  const re = /(^|[^\w.])print[ \t]*\(/g;
  let m;
  while ((m = re.exec(src)) !== null) {
    const open = m.index + m[0].length - 1;
    if (mask[m.index + m[1].length]) continue;
    let depth = 0, close = -1;
    for (let i = open; i < src.length; i++) {
      if (mask[i]) continue;
      if (src[i] === '(') depth++;
      else if (src[i] === ')') { depth--; if (depth === 0) { close = i; break; } }
    }
    if (close < 0) continue;
    let st = open + 1, en = close;
    while (st < en && /\s/.test(src[st])) st++;
    while (en > st && /\s/.test(src[en - 1])) en--;
    if (en <= st) continue;
    out.push({ start: st, end: en, close: close + 1 });
    re.lastIndex = close + 1;
  }
  return out;
}
/** 이미 빈칸이 걸쳐 있는 print 는 버튼을 보이지 않는다 */
const nbvOpenPrints = (src, blanks, lang) =>
  nbvFindPrints(src, lang).filter(p => !(blanks || []).some(b => b.start < p.end && b.end > p.start));

/** root 안의 텍스트 노드를 이어붙인 글자 기준으로 blanks[k] 범위를 make(k, 답) 이 돌려주는 노드로 바꾼다.
 *  표 셀처럼 여러 텍스트 노드에 걸쳐도 처리한다. */
function nbvReplaceRanges(root, blanks, make) {
  const nodes = [];
  const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let pos = 0, n;
  while ((n = w.nextNode())) { nodes.push({ n, s: pos, e: pos + n.nodeValue.length }); pos += n.nodeValue.length; }
  // 노드를 바꾸기 전에 모든 빈칸의 정보를 원본 기준으로 계산해 둔다
  const plan = blanks.map((b, k) => {
    const hit = nodes.filter(x => x.e > b.start && x.s < b.end);
    if (!hit.length) return null;
    const first = hit[0], last = hit[hit.length - 1];
    return {
      k, hit, first, last,
      answer: hit.map(x => x.n.nodeValue.slice(Math.max(b.start, x.s) - x.s, Math.min(b.end, x.e) - x.s)).join(''),
      before: first.n.nodeValue.slice(0, b.start - first.s),
      after: last.n.nodeValue.slice(b.end - last.s),
    };
  });
  // 뒤쪽 빈칸부터 처리해야 앞쪽 빈칸의 위치가 밀리지 않는다. 노드는 교체하지 않고 제자리에서 줄인다.
  for (let i = plan.length - 1; i >= 0; i--) {
    const p = plan[i];
    if (!p) continue;
    const rep = make(p.k, p.answer);
    if (p.first === p.last) {
      p.first.n.nodeValue = p.before;
      const after = document.createTextNode(p.after);
      p.first.n.after(rep, after);
    } else {
      p.first.n.nodeValue = p.before;
      p.first.n.after(rep);
      p.hit.slice(1, -1).forEach(x => { x.n.nodeValue = ''; });
      p.last.n.nodeValue = p.after;
    }
  }
}
const nbvHtmlNode = html => { const t = document.createElement('template'); t.innerHTML = html; return t.content; };
function nbvMarkEl(k, ans) {
  const m = document.createElement('mark');
  m.className = 'nbe-mark';
  m.dataset.b = k;
  m.textContent = ans;
  return m;
}
function nbvBtn(act, label, title) {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'nbe-btn';
  b.dataset.act = act;
  b.textContent = label;
  if (title) b.title = title;
  return b;
}

/* ────────────────────────────────────────────────────────────
   5. 출력 렌더링
   ──────────────────────────────────────────────────────────── */
/** 출력 하나. ctx 가 있으면 빈칸을 입력창(study) 또는 표시(edit)로 그린다. base 는 이 출력의 첫 빈칸 번호 */
function nbvOutputRowEl(o, ctx, base, ci, oi) {
  const edit = !!ctx && ctx.mode === 'edit';
  const make = ctx && ctx.makeTag;
  const blanks = o.blanks || [];
  const tag = (k, ans) => make(base + k, ans);
  const row = document.createElement('div');
  row.className = 'nbv-out nbv-out-' + o.kind + (o.name === 'stderr' ? ' is-err' : '');
  const mark = el => { el.classList.add('nbe-src'); el.dataset.ci = ci; el.dataset.oi = oi; };

  if (o.kind === 'text' || o.kind === 'error') {
    const pre = document.createElement('pre');
    pre.className = 'nbv-pre';
    if (edit) { mark(pre); pre.innerHTML = nbvRawWithMarks({ source: o.text, blanks }) || ' '; }
    else if (blanks.length && make) pre.innerHTML = nbvFill(nbvEsc(nbvTokenized({ source: o.text, blanks }, base)), { source: o.text, blanks }, base, tag);
    else pre.textContent = o.text;
    row.appendChild(pre);
  } else if (o.kind === 'html') {
    const box = document.createElement('div');
    box.className = 'nbv-html';
    box.appendChild(nbvSanitize(o.html));
    if (edit) { mark(box); if (blanks.length) nbvReplaceRanges(box, blanks, nbvMarkEl); }
    else if (blanks.length && make) nbvReplaceRanges(box, blanks, (k, ans) => nbvHtmlNode(tag(k, ans)));
    row.appendChild(box);
  } else if (o.kind === 'md') {
    const box = document.createElement('div');
    if (edit) { box.className = 'nbe-mdsrc'; mark(box); box.innerHTML = nbvRawWithMarks({ source: o.text, blanks }) || ' '; }
    else {
      box.className = 'md-body';
      box.innerHTML = blanks.length && make
        ? nbvFill(nbvMd(nbvTokenized({ source: o.text, blanks }, base)), { source: o.text, blanks }, base, tag)
        : nbvMd(o.text);
    }
    row.appendChild(box);
  } else if (o.kind === 'image' || o.kind === 'svg') {
    const img = document.createElement('img');
    img.className = 'nbv-img';
    img.alt = '출력 이미지';
    img.loading = 'lazy';
    // SVG 는 <img> 로만 그린다 — 이 경로에서는 안의 스크립트가 실행되지 않는다
    img.src = o.kind === 'image'
      ? `data:${o.mime};base64,${o.data}`
      : 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(o.text)));
    row.appendChild(img);
  }
  return row;
}

function nbvOutputsEl(outputs, ctx, base, ci) {
  const edit = !!ctx && ctx.mode === 'edit';
  const wrap = document.createElement('div');
  wrap.className = 'nbv-outs';
  let b = base || 0;
  outputs.forEach((o, oi) => {
    const row = nbvOutputRowEl(o, ctx, b, ci, oi);
    if (edit) {
      // 편집용: 출력마다 수정·삭제 버튼과 빈칸 칩 자리를 붙인다
      const box = document.createElement('div');
      box.className = 'nbe-outrow';
      box.dataset.ci = ci;
      box.dataset.oi = oi;
      const tools = document.createElement('div');
      tools.className = 'nbe-outtools';
      const label = document.createElement('span');
      label.className = 'nbe-outlabel';
      label.textContent = { text: o.name === 'stderr' ? '출력(stderr)' : '출력', error: '오류', html: '표/HTML', md: '마크다운 출력', image: '이미지', svg: 'SVG' }[o.kind] || '출력';
      tools.appendChild(label);
      if (nbvOutEditable(o)) tools.appendChild(nbvBtn('edit-out', '수정'));
      tools.appendChild(nbvBtn('del-out', '삭제'));
      const chips = document.createElement('div');
      chips.className = 'nbe-chips';
      chips.dataset.ci = ci;
      chips.dataset.oi = oi;
      box.append(tools, row);
      if (nbvOutBlankable(o)) box.appendChild(chips);
      wrap.appendChild(box);
    } else {
      wrap.appendChild(row);
    }
    b += (o.blanks || []).length;
  });
  return wrap;
}

/* ────────────────────────────────────────────────────────────
   6. 뷰 빌더 — 학습(빈칸 입력)과 편집(빈칸 지정) 공용
   ──────────────────────────────────────────────────────────── */
function nbvToken(i) {
  let s = 'QQBLANK', n = i;
  do { s += String.fromCharCode(65 + (n % 26)); n = Math.floor(n / 26); } while (n > 0);
  return s + 'ENDQQ';
}
/** 소스에서 빈칸 범위를 토큰으로 치환 — 렌더러(마크다운/하이라이트)를 거친 뒤 토큰 자리에 실제 요소를 끼운다 */
function nbvTokenized(cell, base) {
  let out = '', pos = 0;
  (cell.blanks || []).forEach((b, k) => { out += cell.source.slice(pos, b.start) + nbvToken(base + k); pos = b.end; });
  return out + cell.source.slice(pos);
}
function nbvFill(html, cell, base, make) {
  (cell.blanks || []).forEach((b, k) => { html = html.split(nbvToken(base + k)).join(make(k, cell.source.slice(b.start, b.end))); });
  return html;
}
/** 편집용: 빈칸 토큰에 더해, print 호출 바로 뒤에 버튼 자리(토큰)를 끼워 넣는다 */
function nbvTokenizedEdit(cell, base, marks) {
  const ev = [];
  (cell.blanks || []).forEach((b, k) => ev.push({ pos: b.start, end: b.end, tok: nbvToken(base + k) }));
  marks.forEach(m => ev.push({ pos: m.pos, end: m.pos, tok: m.tok }));
  ev.sort((a, b) => a.pos - b.pos || a.end - b.end);
  let out = '', p = 0;
  ev.forEach(e => { out += cell.source.slice(p, Math.max(p, e.pos)) + e.tok; p = Math.max(p, e.end); });
  return out + cell.source.slice(p);
}
const nbvPrintBtnHtml = p => `<button type="button" class="nbe-pbtn" data-act="print-blank" data-s="${p.start}" data-e="${p.end}" tabindex="-1" title="print 안의 내용을 빈칸으로" aria-label="print 안의 내용을 빈칸으로"></button>`;
const nbvMarkHtml = (k, ans) => `<mark class="nbe-mark" data-b="${k}">${nbvEsc(ans)}</mark>`;

/** 편집용 마크다운: 원문 그대로 보이되 빈칸에 표시를 한다 (렌더링하면 소스 위치를 알 수 없다) */
function nbvRawWithMarks(cell) {
  let out = '', pos = 0;
  (cell.blanks || []).forEach((b, k) => { out += nbvEsc(cell.source.slice(pos, b.start)) + nbvMarkHtml(k, cell.source.slice(b.start, b.end)); pos = b.end; });
  return out + nbvEsc(cell.source.slice(pos));
}

/**
 * ctx.mode === 'study' : ctx.makeTag(idx, answer) → 입력 요소 HTML  (idx 는 카드 전체에서 0부터 이어지는 번호)
 * ctx.mode === 'edit'  : 빈칸을 <mark> 로 표시, 선택 가능한 원문 컨테이너(.nbe-src[data-ci])를 만든다
 */
function nbvCellEl(cell, ci, base, lang, ctx) {
  const edit = ctx.mode === 'edit';
  const el = document.createElement('div');
  el.className = (edit ? 'nbe-cell nbe-' : 'nbv-cell nbv-') + (cell.type === 'code' ? 'code' : 'md');
  el.dataset.ci = ci;
  const label = cell.type === 'code' ? `In [${cell.exec == null ? ' ' : cell.exec}]:` : '';

  if (edit) {
    const head = document.createElement('div');
    head.className = 'nbe-head';
    const lab = document.createElement('span');
    lab.className = 'nbe-label';
    lab.textContent = cell.type === 'code' ? `코드 ${label}` : '마크다운 (원문)';
    const tools = document.createElement('span');
    tools.className = 'nbe-tools';
    tools.append(nbvBtn('edit-cell', '수정', '이 셀의 내용을 직접 고칩니다'), nbvBtn('up', '↑', '위로'), nbvBtn('down', '↓', '아래로'), nbvBtn('del-cell', '삭제'));
    head.append(lab, tools);
    const src = document.createElement(cell.type === 'code' ? 'pre' : 'div');
    src.className = 'nbe-src ' + (cell.type === 'code' ? 'code-surface nbe-codesrc' : 'nbe-mdsrc');
    src.dataset.ci = ci;
    if (cell.type === 'code') {
      const code = document.createElement('code');
      code.className = 'hljs';
      // print(...) 마다 바로 뒤에 '+빈칸' 버튼. 버튼은 글자가 없어서(CSS 로만 그림) 소스 오프셋에 영향이 없다
      const prints = nbvOpenPrints(cell.source, cell.blanks, lang).filter(p => !(cell.blanks || []).some(b => b.start < p.close && p.close < b.end));
      const marks = prints.map((p, i) => ({ pos: p.close, tok: `QQPRINT${i}ENDQQ` }));
      let html = nbvFill(nbvHighlight(nbvTokenizedEdit(cell, base, marks), lang), cell, base, nbvMarkHtml);
      prints.forEach((p, i) => { html = html.split(`QQPRINT${i}ENDQQ`).join(nbvPrintBtnHtml(p)); });
      code.innerHTML = html;
      src.appendChild(code);
    } else {
      src.innerHTML = nbvRawWithMarks(cell);
    }
    const chips = document.createElement('div');
    chips.className = 'nbe-chips';
    chips.dataset.ci = ci;
    el.append(head, src);
    if (cell.type !== 'code' && cell.att) {
      // 원문에는 ![그림](attachment:…) 만 보이므로 어떤 그림인지 작게 보여 준다 (소스 밖이라 오프셋과 무관)
      const strip = document.createElement('div');
      strip.className = 'nbe-att';
      Object.keys(cell.att).forEach(name => {
        const f = cell.att[name];
        if (!NBV_IMG_MIME[f.mime] || !/^[A-Za-z0-9+/=]+$/.test(f.data)) return;
        const img = document.createElement('img');
        img.src = `data:${f.mime};base64,${f.data}`;
        img.alt = img.title = name;
        img.loading = 'lazy';
        strip.appendChild(img);
      });
      el.appendChild(strip);
    }
    el.appendChild(chips);
    if (cell.type === 'code' && (cell.outputs || []).length) {
      const det = document.createElement('details');
      det.className = 'nbe-outs';
      det.open = true;
      const sum = document.createElement('summary');
      sum.textContent = `출력 ${cell.outputs.length}개`;
      det.append(sum, nbvOutputsEl(cell.outputs, ctx, 0, ci));
      el.appendChild(det);
    }
    return el;
  }

  // ── 학습용 ──
  // 비어 있는 셀(편집 중 만들었다가 내용을 안 채운 셀)은 그리지 않는다
  if (!(cell.source || '').trim() && !(cell.blanks || []).length && !(cell.outputs || []).length) return document.createDocumentFragment();
  const gutter = document.createElement('div');
  gutter.className = 'nbv-gutter';
  gutter.textContent = label;
  const body = document.createElement('div');
  body.className = 'nbv-body';
  const make = ctx.makeTag;
  if (cell.type === 'code') {
    const pre = document.createElement('pre');
    pre.className = 'nbv-src code-surface';
    const code = document.createElement('code');
    code.className = 'hljs';
    code.innerHTML = nbvFill(nbvHighlight(nbvTokenized(cell, base), lang), cell, base, (k, ans) => make(base + k, ans));
    pre.appendChild(code);
    body.appendChild(pre);
  } else {
    body.classList.add('md-body');
    nbvAtt = cell.att || null;
    try { body.innerHTML = nbvFill(nbvMd(nbvTokenized(cell, base), lang), cell, base, (k, ans) => make(base + k, ans)); }
    finally { nbvAtt = null; }
  }
  el.append(gutter, body);

  if (cell.type === 'code' && (cell.outputs || []).length) {
    const wrap = document.createElement('div');
    wrap.className = 'nbv-cell nbv-outrow';
    const g = document.createElement('div');
    g.className = 'nbv-gutter nbv-gutter-out';
    const res = cell.outputs.find(o => o.result);
    g.textContent = res && cell.exec != null ? `Out [${cell.exec}]:` : '';
    const b = document.createElement('div');
    b.className = 'nbv-body';
    b.appendChild(nbvOutputsEl(cell.outputs, ctx, base + (cell.blanks || []).length, ci));
    wrap.append(g, b);
    const frag = document.createDocumentFragment();
    frag.append(el, wrap);
    return frag;
  }
  return el;
}

function nbvBuildView(card, ctx) {
  const root = document.createElement('div');
  root.className = 'nbv' + (ctx.mode === 'edit' ? ' nbv-edit' : '') + (ctx.mode !== 'edit' && (card.cells || []).length > 60 ? ' nbv-big' : '');
  const edit = ctx.mode === 'edit';
  const bar = after => {
    const d = document.createElement('div');
    d.className = 'nbe-insert';
    d.dataset.after = after;
    d.append(nbvBtn('add-md', '+ 마크다운'), nbvBtn('add-code', '+ 코드'));
    return d;
  };
  if (edit) root.appendChild(bar(-1));
  let base = 0;
  (card.cells || []).forEach((cell, ci) => {
    root.appendChild(nbvCellEl(cell, ci, base, card.lang, ctx));
    if (edit) root.appendChild(bar(ci));
    base += nbvCellBlankCount(cell);
  });
  return root;
}
/** 셀 ci 앞까지의 빈칸 개수 (소스 + 출력) */
function nbvBaseOf(cells, ci) { let n = 0; for (let k = 0; k < ci; k++) n += nbvCellBlankCount(cells[k]); return n; }

/* ────────────────────────────────────────────────────────────
   7. 선택 영역 ↔ 소스 오프셋 (편집기)
   ──────────────────────────────────────────────────────────── */
function nbvSrcOf(node) {
  const el = node && (node.nodeType === 3 ? node.parentElement : node);
  return el && el.closest ? el.closest('.nbe-src') : null;
}
function nbvOffsetIn(root, node, off) {
  const r = document.createRange();
  r.selectNodeContents(root);
  r.setEnd(node, off);
  return r.toString().length;
}
/** root 안의 [start,end) 글자 범위를 선택 영역으로 만든다 */
function nbvSelectRange(root, start, end) {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let pos = 0, sNode = null, sOff = 0, eNode = null, eOff = 0, n;
  while ((n = walker.nextNode())) {
    const len = n.nodeValue.length;
    if (!sNode && start < pos + len) { sNode = n; sOff = start - pos; }
    if (end <= pos + len) { eNode = n; eOff = end - pos; break; }
    pos += len;
  }
  if (!sNode || !eNode) return false;
  const r = document.createRange();
  r.setStart(sNode, sOff);
  r.setEnd(eNode, eOff);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  return true;
}
