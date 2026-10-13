/**
 * 단어장 PDF → 단어 카드
 *
 * 1) extractPdf  — pdf.js 로 페이지마다 글자 조각(위치·크기·굵게·기울임)과 글자 색을 읽는다 (브라우저/Node 공용)
 * 2) parseVocab  — 조각을 단 → 줄 → 항목으로 묶어 단어 · 품사 · 한국어 뜻 · 영어 뜻 · 예문을 만든다
 *
 * 기대하는 항목 모양 (Nexus Vocabulary List 형식):
 *   accumulate  v.  축적하다; 모으다 — to gather or increase something over time
 *   Example: Over the last 100,000 years, we Sapiens have certainly **accumulated** enormous power.
 * - 표제어는 굵게, 품사는 기울임, 예문 속 정답 단어는 굵게 표시되어 있다
 * - 표제어가 파란색이면 핵심(core) 단어
 * - 2단 편집, 줄바꿈된 뜻·예문, 여러 페이지에 걸친 항목을 처리한다
 * 예문에 굵은 글자가 없는 PDF 는 표제어와 형태가 비슷한 단어(활용형 포함)를 찾아 빈칸으로 삼는다.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.VocabPDF = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const POS_RE = /^(n|v|vt|vi|adj|adv|prep|conj|pron|int|interj|phr|phrase|idiom|det|aux|num|abbr)\.?(\s*(\/|,|&)\s*(n|v|adj|adv|prep|conj)\.?)*$/i;
  const EXAMPLE_RE = /^\s*(examples?|ex|e\.g)\s*[:.]\s*/i;
  const DASH_RE = /\s[—–]\s|\s-{1,2}\s|^[—–]\s|\s[—–]$/;
  const HANGUL_RE = /[ㄱ-ㆎ가-힣]/;
  const WORD_CH = /[A-Za-z0-9'’\-]/;

  /* ============================================================
     1. pdf.js 추출
     ============================================================ */
  function isBlue(rgb) {
    const [r, g, b] = rgb;
    return b >= 110 && b - r >= 60 && b - g >= 25;
  }
  function colorOf(args) {
    // pdf.js 3.x: Uint8ClampedArray[r,g,b] / 일부 버전: '#rrggbb'
    if (!args) return null;
    const a = args[0] !== undefined && args.length === 1 ? args[0] : args;
    if (typeof a === 'string' && /^#[0-9a-f]{6}$/i.test(a)) {
      return [parseInt(a.slice(1, 3), 16), parseInt(a.slice(3, 5), 16), parseInt(a.slice(5, 7), 16)];
    }
    if (a && a.length >= 3) return [a[0], a[1], a[2]].map(Number);
    return null;
  }

  async function extractPdf(data, pdfjsLib, onProgress) {
    // CVE-2024-4367: 글꼴 경로를 eval 로 만들지 않게 한다
    const task = pdfjsLib.getDocument({ data, isEvalSupported: false, disableFontFace: true, useSystemFonts: false });
    const doc = await task.promise;
    let title = '';
    try {
      const meta = await doc.getMetadata();
      title = (meta && meta.info && meta.info.Title) || '';
    } catch (e) { /* 메타데이터 없음 */ }
    const OPS = pdfjsLib.OPS;
    const pages = [];
    for (let p = 1; p <= doc.numPages; p++) {
      const page = await doc.getPage(p);
      const vp = page.getViewport({ scale: 1 });
      const ops = await page.getOperatorList();       // 글꼴 객체를 commonObjs 에 올리고 색 정보를 준다
      const tc = await page.getTextContent();
      const fonts = {};
      const fontOf = name => {
        if (name in fonts) return fonts[name];
        let real = '';
        try { const f = page.commonObjs.get(name); real = (f && (f.name || f.loadedName)) || ''; } catch (e) { /* 아직 없음 */ }
        const st = (tc.styles && tc.styles[name]) || {};
        const id = real + ' ' + (st.fontFamily || '');
        return (fonts[name] = {
          bold: /bold|black|heavy|semibold|demi/i.test(id),
          italic: /italic|oblique/i.test(id),
        });
      };
      const items = [];
      for (const it of tc.items) {
        if (!('str' in it) || !it.str) continue;
        const t = it.transform;
        const f = fontOf(it.fontName);
        items.push({
          str: it.str, x: t[4], y: t[5], w: it.width,
          size: Math.hypot(t[2], t[3]) || Math.abs(t[3]) || 10,
          bold: f.bold, italic: f.italic,
        });
      }
      // 파란 글자(핵심 단어 표시) 모으기 — 같은 색으로 이어진 글자를 한 덩어리로
      const blue = [];
      let cur = null, buf = '';
      const flush = () => { if (buf.trim()) blue.push(buf.replace(/\s+/g, ' ').trim()); buf = ''; };
      for (let i = 0; i < ops.fnArray.length; i++) {
        const fn = ops.fnArray[i], a = ops.argsArray[i];
        if (fn === OPS.setFillRGBColor || fn === OPS.setFillColor || fn === OPS.setFillGray || fn === OPS.setFillCMYKColor) {
          const c = fn === OPS.setFillGray ? null : colorOf(a);
          const b = !!(c && fn !== OPS.setFillCMYKColor && isBlue(c));
          if (b !== cur) { flush(); cur = b; }
        } else if ((fn === OPS.showText || fn === OPS.showSpacedText) && cur) {
          const glyphs = (a && a[0]) || [];
          for (const g of glyphs) {
            if (g && typeof g === 'object' && g.unicode) buf += g.unicode;
            else if (typeof g === 'number' && g < -200) buf += ' ';
          }
        } else if (fn === OPS.endText && cur) { buf += ' '; }
      }
      flush();
      pages.push({ width: vp.width, height: vp.height, items, blue });
      page.cleanup();
      if (onProgress) onProgress(p, doc.numPages);
    }
    await doc.destroy();
    return { title, pages };
  }

  /* ============================================================
     2. 조각 → 줄
     ============================================================ */
  function buildLines(items) {
    const sorted = items.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
    const lines = [];
    for (const it of sorted) {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.y - it.y) <= Math.max(1.5, Math.min(last.size, it.size) * 0.45)) last.items.push(it);
      else lines.push({ y: it.y, size: it.size, items: [it] });
      const l = lines[lines.length - 1];
      l.size = Math.max(l.size, it.size);
    }
    for (const l of lines) {
      l.items.sort((a, b) => a.x - b.x);
      const segs = [];
      let prev = null;
      for (const it of l.items) {
        if (prev) {
          const gap = it.x - (prev.x + prev.w);
          const lastCh = segs.length ? segs[segs.length - 1].text.slice(-1) : '';
          if (gap > l.size * 0.18 && !/\s/.test(lastCh) && !/^\s/.test(it.str)) {
            segs.push({ text: ' ', bold: prev.bold && it.bold, italic: false });
          }
        }
        segs.push({ text: it.str, bold: it.bold, italic: it.italic });
        prev = it;
      }
      l.segs = segs;
      l.x = l.items[0].x;
      l.text = segs.map(s => s.text).join('');
    }
    return lines;
  }

  /** 페이지마다 반복되는 머리말·꼬리말(쪽 번호만 다른 줄)을 찾는다 */
  function repeatedKeys(pageLines) {
    if (pageLines.length < 2) return new Set();
    const count = new Map();
    for (const lines of pageLines) {
      const seen = new Set();
      for (const l of lines) {
        const k = l.text.replace(/\d+/g, '#').replace(/\s+/g, ' ').trim();
        if (k && !seen.has(k)) { seen.add(k); count.set(k, (count.get(k) || 0) + 1); }
      }
    }
    const out = new Set();
    for (const [k, n] of count) if (n >= Math.max(2, Math.ceil(pageLines.length * 0.6))) out.add(k);
    return out;
  }

  /** 가운데 세로 여백을 찾아 단을 나눈다 — 한 단짜리 페이지면 [] */
  function columnSplit(items, width) {
    const mid = width / 2;
    const lo = width * 0.3, hi = width * 0.7;
    // 각 x 위치가 조각에 덮이는지 — 덮이지 않는 가장 넓은 구간이 단 사이 여백
    const step = 2, n = Math.ceil(width / step);
    const cover = new Uint16Array(n);
    for (const it of items) {
      if (!it.str.trim()) continue;
      const a = Math.max(0, Math.floor(it.x / step)), b = Math.min(n - 1, Math.ceil((it.x + it.w) / step));
      for (let i = a; i <= b; i++) cover[i]++;
    }
    let best = null, run = null;
    for (let i = Math.floor(lo / step); i <= Math.ceil(hi / step); i++) {
      if (cover[i] === 0) { if (!run) run = { a: i, b: i }; else run.b = i; }
      else if (run) { if (!best || run.b - run.a > best.b - best.a) best = run; run = null; }
    }
    if (run && (!best || run.b - run.a > best.b - best.a)) best = run;
    if (best && (best.b - best.a) * step >= 8) return [((best.a + best.b) / 2) * step];
    // 여백이 안 보이면(긴 제목 등) 왼쪽 단에서 시작하는 줄이 오른쪽 반까지 넘어가는지로 판단
    const rightStarts = items.filter(it => it.str.trim() && it.x > mid && it.x < hi + width * 0.1).length;
    return rightStarts > items.length * 0.15 ? [mid] : [];
  }

  /* ============================================================
     3. 줄 → 항목
     ============================================================ */
  function isEntryStart(line) {
    const segs = line.segs.filter(s => s.text.trim());
    if (!segs.length || !segs[0].bold) return false;
    if (EXAMPLE_RE.test(line.text)) return false;
    if (!/^[A-Za-z(]/.test(segs[0].text.trim())) return false;
    // 굵은 표제어 바로 뒤(세 조각 이내)에 품사가 와야 한다
    for (let i = 1; i < Math.min(segs.length, 4); i++) {
      const t = segs[i].text.trim();
      if (segs[i].bold) continue;
      return POS_RE.test(t) || /^(n|v|adj|adv|prep|conj|pron|phr)\.(\s|$)/i.test(t);
    }
    return false;
  }

  /** segs 를 하나의 문자열로 — 굵은 범위를 함께 돌려준다. 줄 사이는 공백으로 잇는다 */
  function joinSegs(lines) {
    let text = '';
    const bold = [];
    lines.forEach((segs, li) => {
      if (li > 0 && text && !/[\s\-‐]$/.test(text)) text += ' ';
      else if (li > 0 && /[A-Za-z]-$/.test(text)) { /* 하이픈으로 끝난 줄은 그대로 잇는다 */ }
      for (const s of segs) {
        const start = text.length;
        text += s.text;
        if (s.bold && s.text.trim()) bold.push([start, text.length]);
      }
    });
    // 정리: 연속 공백 축소 (범위 보정)
    let out = '', map = [];
    for (let i = 0; i < text.length; i++) {
      const c = /\s/.test(text[i]) ? ' ' : text[i];
      if (c === ' ' && (out.endsWith(' ') || !out.length)) { map.push(out.length); continue; }
      map.push(out.length);
      out += c;
    }
    map.push(out.length);
    const trimmedEnd = out.replace(/\s+$/, '').length;
    out = out.slice(0, trimmedEnd);
    const ranges = bold.map(([a, b]) => [map[a], Math.min(map[b], trimmedEnd)]).filter(([a, b]) => b > a);
    return { text: out, bold: mergeRanges(out, ranges) };
  }
  function mergeRanges(text, ranges) {
    ranges.sort((a, b) => a[0] - b[0]);
    const out = [];
    for (const r of ranges) {
      const last = out[out.length - 1];
      if (last && (r[0] <= last[1] || !text.slice(last[1], r[0]).trim())) last[1] = Math.max(last[1], r[1]);
      else out.push([r[0], r[1]]);
    }
    // 범위 끝의 공백·문장부호는 빼고, 단어 중간에서 잘렸으면 단어 경계까지 넓힌다
    return out.map(([a, b]) => {
      while (a < b && /[\s"“”'‘’(]/.test(text[a])) a++;
      while (b > a && /[\s"“”'‘’),.;:!?]/.test(text[b - 1])) b--;
      // 소유격 's 는 넓히지 않는다 (Sorcerer’s → Sorcerer), 하이픈 낱말은 통째로 (anti-vaxxers)
      const AL = /[A-Za-z0-9]/;
      while (a > 0 && (AL.test(text[a - 1]) || (text[a - 1] === '-' && AL.test(text[a - 2] || '')))) a--;
      while (b < text.length && (AL.test(text[b]) || (text[b] === '-' && AL.test(text[b + 1] || '')))) b++;
      return [a, b];
    }).filter(([a, b]) => b > a);
  }

  function parseEntry(headLines, exampleBlocks) {
    // 머리 줄: 표제어(굵게) · 품사(기울임) · 한국어 뜻 — 영어 뜻
    const segs = [];
    headLines.forEach((l, i) => {
      const prev = segs.length ? segs[segs.length - 1].text : '';
      if (i > 0 && !/[A-Za-z]-$/.test(prev)) segs.push({ text: ' ' });    // 하이픈에서 줄바꿈된 낱말은 붙인다
      segs.push(...l);
    });
    let i = 0;
    while (i < segs.length && !segs[i].text.trim()) i++;
    let term = '';
    while (i < segs.length && (segs[i].bold || (!segs[i].text.trim() && segs[i + 1] && segs[i + 1].bold))) term += segs[i++].text;
    let pos = '';
    while (i < segs.length && !segs[i].text.trim()) i++;
    if (i < segs.length && POS_RE.test(segs[i].text.trim().replace(/\s+/g, ' '))) pos = segs[i++].text.trim();
    let rest = segs.slice(i).map(s => s.text).join('').replace(/\s+/g, ' ').trim();
    if (!pos) {
      const m = /^((?:n|v|adj|adv|prep|conj|pron|phr)\.)\s*/i.exec(rest);
      if (m) { pos = m[1]; rest = rest.slice(m[0].length); }
    }
    let ko = '', def = '';
    const dm = DASH_RE.exec(rest);
    if (dm) { ko = rest.slice(0, dm.index).trim(); def = rest.slice(dm.index + dm[0].length).trim(); }
    else if (HANGUL_RE.test(rest)) {
      // 대시가 없으면 한글이 끝나는 곳에서 자른다
      const m = /^([^A-Za-z]*[ㄱ-ㆎ가-힣][^A-Za-z]*)(.*)$/.exec(rest);
      if (m) { ko = m[1].trim(); def = m[2].trim(); } else def = rest;
    } else def = rest;
    ko = ko.replace(/^[—–\-\s]+|[—–\-\s]+$/g, '');
    def = def.replace(/^[—–\-\s]+/, '');
    term = term.replace(/\s+/g, ' ').trim();

    const examples = exampleBlocks.map(lines => {
      const j = joinSegs(lines);
      const m = EXAMPLE_RE.exec(j.text);
      const cut = m ? m[0].length : 0;
      const text = j.text.slice(cut).trim();
      const lead = j.text.slice(cut).length - j.text.slice(cut).trimStart().length + cut;
      let blanks = j.bold.map(([a, b]) => [a - lead, b - lead]).filter(([a, b]) => a >= 0 && b <= text.length && b > a);
      blanks = blanks.filter(([a, b]) => looksLikeTarget(text.slice(a, b), term));
      if (!blanks.length) blanks = findTarget(text, term);
      return { text, blanks };
    }).filter(e => e.text);
    return { term, pos, ko, def, examples };
  }

  /** 굵게 표시된 부분이 실제로 그 단어인지 (예문 속 다른 강조와 구별) */
  function looksLikeTarget(s, term) {
    const a = baseForms(term), t = s.toLowerCase();
    return a.some(f => similar(t, f)) || t.split(/\s+/).some(w => a.some(f => similar(w, f)));
  }

  /* ============================================================
     4. 활용형 찾기 (굵은 표시가 없을 때)
     ============================================================ */
  /** 'aberration(s)' → ['aberration', 'aberrations'] */
  function baseForms(term) {
    const t = String(term || '').toLowerCase().replace(/[’]/g, "'").trim();
    const out = new Set();
    const m = /^(.*?)\(([^)]*)\)(.*)$/.exec(t);
    if (m) { out.add((m[1] + m[3]).trim()); out.add((m[1] + m[2] + m[3]).trim()); }
    else out.add(t);
    t.split('/').forEach(x => { if (x.trim() && x !== t) out.add(x.trim()); });
    return [...out].filter(Boolean);
  }
  function lcp(a, b) { let i = 0; while (i < a.length && i < b.length && a[i] === b[i]) i++; return i; }
  /** 표제어 h 와 문장 속 단어 w 가 같은 단어의 활용형으로 보이는지 */
  function similar(w, h) {
    w = w.toLowerCase().replace(/[’]/g, "'").replace(/^[^a-z0-9]+|[^a-z0-9]+$/g, '');
    h = h.toLowerCase();
    if (!w || !h) return false;
    if (w === h) return true;
    if (h.length <= 3) return w === h + 's' || w === h + 'es' || w === h + 'ed' || w === h + 'ing';
    const p = lcp(w, h);
    const stem = h.replace(/(e|y|ie)$/, '');
    if (p >= stem.length && w.length - p <= 5) return true;          // accumulate → accumulated, apply → applied
    if (p >= Math.max(4, h.length - 2) && w.length - p <= 5) return true;  // 끝 자음 중복·변형 (admit → admitted)
    return false;
  }
  /** 문장에서 표제어(여러 단어 가능)의 위치를 찾는다 → [[start, end], ...] */
  function findTarget(text, term) {
    const forms = baseForms(term);
    const toks = [];
    const re = /[A-Za-z0-9][A-Za-z0-9'’\-]*/g;
    let m;
    while ((m = re.exec(text))) toks.push({ s: m.index, e: m.index + m[0].length, w: m[0] });
    for (const f of forms) {
      const words = f.split(/\s+/).filter(Boolean);
      const hits = [];
      for (let i = 0; i + words.length <= toks.length; i++) {
        let ok = true;
        for (let k = 0; k < words.length && ok; k++) ok = similar(toks[i + k].w, words[k]);
        if (ok) { hits.push([toks[i].s, toks[i + words.length - 1].e]); i += words.length - 1; }
      }
      if (hits.length) return hits;
    }
    return [];
  }

  /* ============================================================
     5. 전체 파싱
     ============================================================ */
  function parseVocab(doc) {
    const stream = [];        // 읽는 순서대로: 페이지 → 단 → 줄
    const pageLines = doc.pages.map(pg => {
      const splits = columnSplit(pg.items, pg.width);
      const cols = [[]];
      splits.forEach(() => cols.push([]));
      for (const it of pg.items) {
        let c = 0;
        while (c < splits.length && it.x >= splits[c]) c++;
        cols[c].push(it);
      }
      return cols.map(items => buildLines(items));
    });
    const rep = repeatedKeys(pageLines.map(cols => cols.flat()));
    pageLines.forEach((cols, pi) => {
      const pg = doc.pages[pi];
      cols.forEach(lines => {
        for (const l of lines) {
          const k = l.text.replace(/\d+/g, '#').replace(/\s+/g, ' ').trim();
          if (!k || rep.has(k)) continue;
          if (/^#$/.test(k) && (l.y < pg.height * 0.08 || l.y > pg.height * 0.92)) continue;  // 쪽 번호만 있는 줄
          stream.push(l);
        }
      });
    });

    const entries = [];
    let cur = null;
    const close = () => {
      if (!cur) return;
      const e = parseEntry(cur.head, cur.examples);
      if (e.term) entries.push(e);
      cur = null;
    };
    for (const l of stream) {
      if (isEntryStart(l)) { close(); cur = { head: [l.segs], examples: [] }; continue; }
      if (!cur) continue;          // 첫 항목 이전(제목·범례)은 버린다
      if (EXAMPLE_RE.test(l.text)) { cur.examples.push([l.segs]); continue; }
      if (cur.examples.length) cur.examples[cur.examples.length - 1].push(l.segs);
      else cur.head.push(l.segs);
    }
    close();

    // 핵심 단어 = 파란색 표제어
    const blue = new Set();
    doc.pages.forEach(pg => (pg.blue || []).forEach(t => blue.add(t.toLowerCase())));
    const blueList = [...blue];
    for (const e of entries) {
      const t = e.term.toLowerCase();
      e.core = blue.has(t) || blueList.some(b => b.startsWith(t + ' ') && b.length < t.length + 4);
    }
    return { title: doc.title || '', entries, hasCore: entries.some(e => e.core) };
  }

  /** 뜻풀이 안에 정답이 그대로 들어 있으면 가린다 (예: disposal — "... at our disposal") */
  function maskInText(text, term, answers) {
    let out = String(text || '');
    const words = new Set();
    baseForms(term).forEach(f => words.add(f));
    (answers || []).forEach(a => a && words.add(String(a).toLowerCase()));
    const ranges = [];
    for (const w of words) {
      findTarget(out, w).forEach(r => ranges.push(r));
    }
    ranges.sort((a, b) => b[0] - a[0]);
    let lastStart = Infinity;
    for (const [a, b] of ranges) {
      if (b > lastStart) continue;
      out = out.slice(0, a) + '____' + out.slice(b);
      lastStart = a;
    }
    return out;
  }

  return { extractPdf, parseVocab, findTarget, baseForms, similar, maskInText, buildLines, columnSplit };
});
