/* ============================================================
   docx.js — Word(.docx) 문서 → 노트북 뷰어용 마크다운 셀

   Google Docs 는 '파일 → 다운로드 → Microsoft Word(.docx)' 로 받으면 된다.
   문단 · 목록(단계) · 표 · 그림을 마크다운 셀로 옮기고, 그림은 셀 첨부(att)로 넣는다.
   결과는 .ipynb 를 읽었을 때와 같은 { lang, cells } 라서 뷰어 · 빈칸 · 편집을 그대로 쓴다.

   - .docx 는 ZIP 이다. 외부 라이브러리 없이 브라우저 내장 DecompressionStream 으로 푼다.
   - 문서는 외부에서 온 파일이다. 글자는 마크다운 원문으로만 옮기고(렌더링 때 전부 이스케이프),
     그림은 png/jpeg/gif/webp 의 base64 만 받는다.
   ============================================================ */
'use strict';

const DX_MAX_XML = 20 * 1048576;     // 압축을 푼 XML 한 개 상한
const DX_MAX_MEDIA = 8 * 1048576;    // 그림 한 장 원본 상한 (이보다 크면 건너뜀)

/* ── 1. ZIP ─────────────────────────────────────────────── */
function dxZip(buf) {
  const u8 = new Uint8Array(buf), dv = new DataView(buf);
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('ZIP 형식이 아닙니다');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const entries = new Map();
  const dec = new TextDecoder();
  for (let k = 0; k < count; k++) {
    if (p + 46 > u8.length || dv.getUint32(p, true) !== 0x02014b50) throw new Error('ZIP 목록이 손상되었습니다');
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    entries.set(dec.decode(u8.subarray(p + 46, p + 46 + nlen)), {
      method: dv.getUint16(p + 10, true),
      csize: dv.getUint32(p + 20, true),
      usize: dv.getUint32(p + 24, true),
      off: dv.getUint32(p + 42, true),
    });
    p += 46 + nlen + xlen + clen;
  }
  async function bytes(name, max) {
    const e = entries.get(name);
    if (!e) return null;
    if (e.usize > max) return null;                         // 압축 폭탄·지나치게 큰 항목은 읽지 않는다
    const lh = e.off;
    if (dv.getUint32(lh, true) !== 0x04034b50) throw new Error('ZIP 항목이 손상되었습니다');
    const start = lh + 30 + dv.getUint16(lh + 26, true) + dv.getUint16(lh + 28, true);
    const data = u8.subarray(start, start + e.csize);
    if (e.method === 0) return data;
    if (e.method !== 8) throw new Error('지원하지 않는 압축 방식입니다');
    const out = await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer();
    return new Uint8Array(out);
  }
  return {
    has: n => entries.has(n),
    bytes,
    async text(n) { const b = await bytes(n, DX_MAX_XML); return b ? new TextDecoder().decode(b) : null; },
  };
}

/* ── 2. XML 도우미 ──────────────────────────────────────── */
const dxXml = t => {
  const d = new DOMParser().parseFromString(t, 'application/xml');
  if (d.getElementsByTagName('parsererror').length) throw new Error('문서 XML 을 읽을 수 없습니다');
  return d;
};
const dxKids = (el, name) => el ? [...el.children].filter(c => c.localName === name) : [];
const dxKid = (el, name) => dxKids(el, name)[0] || null;
const dxAll = (el, name) => el ? [...el.getElementsByTagName('*')].filter(c => c.localName === name) : [];
const dxVal = (el, attr = 'w:val') => el ? el.getAttribute(attr) : null;
const dxOn = el => !!el && !/^(0|false|none)$/i.test(dxVal(el) || '');   // <w:b/> 또는 <w:b w:val="1"/>

function dxRels(doc) {
  const map = new Map();
  if (doc) dxAll(doc.documentElement, 'Relationship').forEach(r => map.set(r.getAttribute('Id'), { target: r.getAttribute('Target') || '', external: r.getAttribute('TargetMode') === 'External' }));
  return map;
}
/** numId → ilvl → 번호 형식('bullet' | 'decimal' …) */
function dxNumbering(doc) {
  const abs = new Map(), nums = new Map();
  if (!doc) return nums;
  dxKids(doc.documentElement, 'abstractNum').forEach(a => {
    const lv = new Map();
    dxKids(a, 'lvl').forEach(l => lv.set(Number(dxVal(l, 'w:ilvl')), dxVal(dxKid(l, 'numFmt')) || 'bullet'));
    abs.set(dxVal(a, 'w:abstractNumId'), lv);
  });
  dxKids(doc.documentElement, 'num').forEach(n => nums.set(dxVal(n, 'w:numId'), abs.get(dxVal(dxKid(n, 'abstractNumId'))) || new Map()));
  return nums;
}
/** 스타일 id → 제목 단계 (Title=0, Heading1=1 …) */
function dxHeadingStyles(doc) {
  const map = new Map();
  if (!doc) return map;
  dxKids(doc.documentElement, 'style').forEach(st => {
    const id = dxVal(st, 'w:styleId');
    const name = (dxVal(dxKid(st, 'name')) || '').toLowerCase();
    const ol = dxVal(dxKid(dxKid(st, 'pPr'), 'outlineLvl'));
    let m;
    if (name === 'title') map.set(id, 0);
    else if ((m = name.match(/^heading\s*(\d)$/))) map.set(id, Number(m[1]));
    else if (ol != null && Number(ol) < 6) map.set(id, Number(ol) + 1);
  });
  return map;
}

/* ── 3. 문단 → 마크다운 인라인 ─────────────────────────── */
const DX_MONO = /consolas|courier|mono|menlo|source code|d2coding|fira code|jetbrains/i;

/** 문단 안의 글자 조각을 서식과 함께 모은다 */
function dxSegments(p, ctx) {
  const segs = [];
  const run = (r, link) => {
    const rPr = dxKid(r, 'rPr');
    const fonts = dxKid(rPr, 'rFonts');
    const fmt = {
      b: dxOn(dxKid(rPr, 'b')),
      i: dxOn(dxKid(rPr, 'i')),
      s: dxOn(dxKid(rPr, 'strike')) || dxOn(dxKid(rPr, 'dstrike')),
      code: !!fonts && DX_MONO.test((fonts.getAttribute('w:ascii') || '') + ' ' + (fonts.getAttribute('w:hAnsi') || '')),
      link,
    };
    [...r.children].forEach(c => {
      const n = c.localName;
      if (n === 't') segs.push({ ...fmt, text: c.textContent });
      else if (n === 'tab') segs.push({ ...fmt, text: '\t' });
      else if (n === 'br' || n === 'cr') { if (dxVal(c, 'w:type') !== 'page') segs.push({ text: '\n' }); }
      else if (n === 'noBreakHyphen') segs.push({ ...fmt, text: '-' });
      else if (n === 'drawing' || n === 'pict' || n === 'object') {
        const blip = dxAll(c, 'blip')[0], vimg = dxAll(c, 'imagedata')[0];
        const rid = blip ? (blip.getAttribute('r:embed') || blip.getAttribute('r:link')) : vimg ? vimg.getAttribute('r:id') : null;
        if (rid) segs.push({ img: rid });
      }
    });
  };
  const walk = (el, link) => {
    [...el.children].forEach(c => {
      const n = c.localName;
      if (n === 'r') run(c, link);
      else if (n === 'hyperlink') {
        const rel = ctx.rels.get(c.getAttribute('r:id'));
        walk(c, rel && rel.external && /^https?:\/\//i.test(rel.target) ? rel.target : link);
      }
      else if (n === 'ins' || n === 'smartTag' || n === 'customXml' || n === 'fldSimple') walk(c, link);
      else if (n === 'sdt') walk(dxKid(c, 'sdtContent') || c, link);
      // del(삭제 추적)·pPr 등은 건너뛴다
    });
  };
  walk(p, null);
  return segs;
}

/** 서식이 같은 조각을 합쳐 마크다운으로. 표시(** 등)는 공백 밖으로 빼서 붙인다 */
function dxInline(segs) {
  const same = (a, b) => a.b === b.b && a.i === b.i && a.s === b.s && a.code === b.code && a.link === b.link;
  const merged = [];
  segs.forEach(s => {
    if (s.img || s.text === '\n') { merged.push(s); return; }
    const last = merged[merged.length - 1];
    if (last && !last.img && last.text !== '\n' && same(last, s)) last.text += s.text;
    else merged.push({ ...s });
  });
  return merged.map(s => {
    if (s.img) return '';
    if (s.text === '\n') return '<br>';
    let t = s.text.replace(/\t/g, ' ').replace(/ /g, ' ');
    const lead = t.match(/^\s*/)[0], tail = t.slice(lead.length).match(/\s*$/)[0];
    let core = t.slice(lead.length, t.length - tail.length);
    if (!core) return t;
    if (s.code) core = '`' + core.replace(/`/g, "'") + '`';
    else {
      if (s.s) core = '~~' + core + '~~';
      if (s.i) core = '*' + core + '*';
      if (s.b) core = '**' + core + '**';
    }
    if (s.link) core = '[' + core + '](' + s.link.replace(/[()\s]/g, encodeURIComponent) + ')';
    return lead + core + tail;
  }).join('');
}

/* ── 4. 그림 ────────────────────────────────────────────── */
const DX_IMG_EXT = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' };
function dxB64(u8) {
  let s = '';
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
  return btoa(s);
}
/** 저장 한도(NBV_MAX_IMG)보다 크면 줄여서 JPEG 로 다시 굽는다 */
async function dxFitImage(u8, mime) {
  let data = dxB64(u8);
  if (data.length <= NBV_MAX_IMG) return { mime, data };
  try {
    const bmp = await createImageBitmap(new Blob([u8], { type: mime }));
    let scale = Math.min(1, 1400 / Math.max(bmp.width, bmp.height));
    for (let k = 0; k < 6; k++) {
      const c = document.createElement('canvas');
      c.width = Math.max(1, Math.round(bmp.width * scale));
      c.height = Math.max(1, Math.round(bmp.height * scale));
      const g = c.getContext('2d');
      g.fillStyle = '#fff';
      g.fillRect(0, 0, c.width, c.height);
      g.drawImage(bmp, 0, 0, c.width, c.height);
      data = c.toDataURL('image/jpeg', 0.82).split(',')[1] || '';
      if (data && data.length <= NBV_MAX_IMG) return { mime: 'image/jpeg', data };
      scale *= 0.75;
    }
  } catch (e) { /* 디코딩 못 하는 그림 */ }
  return null;
}

/* ── 5. 표 → 마크다운 표 ────────────────────────────────── */
function dxTable(tbl, ctx) {
  const rows = dxKids(tbl, 'tr').map(tr => {
    const cells = [];
    dxKids(tr, 'tc').forEach(tc => {
      const tcPr = dxKid(tc, 'tcPr');
      const vm = dxKid(tcPr, 'vMerge');
      const cont = vm && dxVal(vm) !== 'restart';           // 위 칸과 합쳐진 칸
      const text = cont ? '' : dxKids(tc, 'p').map(p => {
        const segs = dxSegments(p, ctx);
        segs.filter(s => s.img).forEach(s => ctx.lostImages++);
        const t = dxInline(segs).trim();
        return t && dxKid(dxKid(p, 'pPr'), 'numPr') ? '• ' + t : t;
      }).filter(Boolean).join('<br>');
      // 표 구분자와 겹치지 않게 | 는 전각 문자로
      cells.push(text.replace(/\|/g, '｜'));
      const span = Number(dxVal(dxKid(tcPr, 'gridSpan')) || 1);
      for (let k = 1; k < span; k++) cells.push('');
    });
    return cells;
  }).filter(r => r.length);
  if (!rows.length) return [];
  const w = Math.max(...rows.map(r => r.length));
  const line = r => '| ' + Array.from({ length: w }, (_, k) => r[k] || ' ').join(' | ') + ' |';
  return [line(rows[0]), '| ' + Array(w).fill('---').join(' | ') + ' |', ...rows.slice(1).map(line)];
}

/* ── 6. 문서 → 셀 ───────────────────────────────────────── */
async function dxParseDocx(buf) {
  if (typeof DecompressionStream === 'undefined') throw new Error('이 브라우저는 .docx 압축 해제를 지원하지 않습니다 (Firefox 113+, Chrome 103+)');
  const zip = dxZip(buf);
  const docXml = await zip.text('word/document.xml');
  if (!docXml) throw new Error('word/document.xml 이 없습니다 — Word 문서가 맞는지 확인해 주세요');
  const doc = dxXml(docXml);
  const opt = async n => { const t = await zip.text(n); return t ? dxXml(t) : null; };
  const ctx = {
    rels: dxRels(await opt('word/_rels/document.xml.rels')),
    nums: dxNumbering(await opt('word/numbering.xml')),
    heads: dxHeadingStyles(await opt('word/styles.xml')),
    lostImages: 0,
  };
  const body = dxAll(doc.documentElement, 'body')[0];
  if (!body) throw new Error('문서 본문이 비어 있습니다');

  // 본문의 블록을 평평하게: 문단 · 표 (콘텐츠 컨트롤 안쪽도 펼친다)
  const blocks = [];
  const collect = el => [...el.children].forEach(c => {
    if (c.localName === 'p' || c.localName === 'tbl') blocks.push(c);
    else if (c.localName === 'sdt') collect(dxKid(c, 'sdtContent') || c);
  });
  collect(body);

  // 블록 → 중간 표현
  const items = blocks.map(b => {
    if (b.localName === 'tbl') return { kind: 'table', lines: dxTable(b, ctx) };
    const pPr = dxKid(b, 'pPr');
    const segs = dxSegments(b, ctx);
    const text = dxInline(segs).replace(/^(<br>)+|(<br>)+$/g, '');
    const numPr = dxKid(pPr, 'numPr');
    const style = dxVal(dxKid(pPr, 'pStyle'));
    const ind = dxKid(pPr, 'ind');
    return {
      kind: 'p',
      text,
      imgs: segs.filter(s => s.img).map(s => s.img),
      head: style != null && ctx.heads.has(style) ? ctx.heads.get(style) : null,
      list: numPr ? { numId: dxVal(dxKid(numPr, 'numId')), lvl: Number(dxVal(dxKid(numPr, 'ilvl')) || 0) } : null,
      indented: !!ind && Number(ind.getAttribute('w:left') || ind.getAttribute('w:start') || 0) > 0,
      tabbed: /^\s/.test(segs.filter(s => !s.img).map(s => s.text || '').join('')),
    };
  });

  // 그림 데이터
  const att = new Map();          // rId → { name, mime, data } | null
  let imgNo = 0;
  for (const it of items) {
    for (const rid of it.imgs || []) {
      if (att.has(rid)) continue;
      const rel = ctx.rels.get(rid);
      let entry = null;
      if (rel && !rel.external) {
        const path = rel.target.startsWith('/') ? rel.target.slice(1) : 'word/' + rel.target.replace(/^\.\//, '');
        const ext = (path.split('.').pop() || '').toLowerCase();
        const mime = DX_IMG_EXT[ext];
        const raw = mime ? await zip.bytes(path, DX_MAX_MEDIA) : null;
        const fit = raw ? await dxFitImage(raw, mime) : null;
        if (fit) entry = { name: `image${++imgNo}.${fit.mime === 'image/jpeg' ? 'jpg' : ext}`, ...fit };
      }
      if (!entry) ctx.lostImages++;
      att.set(rid, entry);
    }
  }

  // 셀로 묶기: 빈 문단(단락 구분)과 제목에서 새 셀을 시작한다
  const cells = [];
  let cur = null;
  const flush = () => {
    if (!cur) return;
    const source = cur.lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
    if (source) {
      const cell = { type: 'markdown', source, blanks: [] };
      if (Object.keys(cur.att).length) cell.att = cur.att;
      cells.push(cell);
    }
    cur = null;
  };
  const open = () => { if (!cur) cur = { lines: [], att: {}, lastList: -1 }; return cur; };
  const counters = new Map();      // numId → [단계별 번호]
  const nextNonEmpty = k => { for (let j = k + 1; j < items.length; j++) { const x = items[j]; if (x.kind === 'table' || x.text.trim() || (x.imgs || []).length) return x; } return null; };

  items.forEach((it, k) => {
    if (it.kind === 'table') {
      const c = open();
      if (it.lines.length) c.lines.push('', ...it.lines, '');
      c.lastList = -1;
      return;
    }
    const hasImg = (it.imgs || []).some(r => att.get(r));
    if (!it.text.trim() && !hasImg) {            // 빈 문단 = 단락 구분
      if (!it.imgs.length) flush();
      return;
    }
    // 제목: 문서 스타일이 제목이거나, 스타일이 없어도 '짧고 콜론·마침표 없이 바로 목록/표가 이어지는 줄'
    let level = it.head;
    // 들여쓰기·탭으로 시작하는 줄은 앞 목록 항목에 딸린 줄(ex) …)이지 제목이 아니다
    if (level == null && !it.list && !it.tabbed && !it.indented && it.text.trim() && it.text.length <= 40 && !/[:：.。!?]/.test(it.text.slice(-1)) && !/[:：`*<]/.test(it.text)) {
      const nx = nextNonEmpty(k);
      if (nx && (nx.kind === 'table' || nx.list)) level = 2;     // ### 로
    }
    if (level != null && it.text.trim()) {
      flush();
      const c = open();
      c.lines.push('#'.repeat(Math.min(Math.max(level + 1, 2), 4)) + ' ' + it.text.replace(/<br>/g, ' ').trim());
      c.lastList = -1;
    } else if (it.list) {
      const c = open();
      const fmt = (ctx.nums.get(it.list.numId) || new Map()).get(it.list.lvl) || 'bullet';
      const cnt = counters.get(it.list.numId) || [];
      cnt[it.list.lvl] = (cnt[it.list.lvl] || 0) + 1;
      cnt.length = it.list.lvl + 1;                 // 더 깊은 단계 번호는 다시 1부터
      counters.set(it.list.numId, cnt);
      const marker = fmt === 'bullet' || fmt === 'none' ? '-' : cnt[it.list.lvl] + '.';
      c.lines.push('  '.repeat(it.list.lvl) + marker + ' ' + it.text.trim());
      c.lastList = it.list.lvl;
    } else if (it.text.trim()) {
      const c = open();
      const last = c.lines.length - 1;
      if (c.lastList >= 0 && (it.tabbed || it.indented) && last >= 0 && c.lines[last].trim() && !/^\s*!\[/.test(c.lines[last])) {
        c.lines[last] += '<br>' + it.text.trim();      // 목록 항목에 딸린 줄 (ex) …) — 같은 항목 안에서 줄만 바꾼다
      } else {
        if (c.lastList >= 0) c.lines.push('');
        c.lines.push(it.text.trim());
        c.lastList = -1;
      }
    }
    // 문단 안의 그림은 그 자리 다음 줄에
    (it.imgs || []).forEach(rid => {
      const a = att.get(rid);
      if (!a) return;
      const c = open();
      c.att[a.name] = { mime: a.mime, data: a.data };
      const pad = c.lastList >= 0 ? '  '.repeat(c.lastList + 1) : '';
      c.lines.push(pad + `![그림 ${a.name.replace(/\D/g, '')}](attachment:${a.name})`);
    });
  });
  flush();
  return { lang: 'text', cells, lostImages: ctx.lostImages, images: imgNo };
}
