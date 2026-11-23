/* Import wyników z PDF Karting Managera do Exceli (generalka + dywizje).
   Działa w przeglądarce: pdf.js (lib/pdf.min.js) + JSZip (lib/jszip.min.js). */
(function () {
'use strict';
var NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
var RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/* ---------- punktacja ---------- */
function pts(p) { return p < 1 ? 0 : p <= 4 ? [40, 37, 33, 31][p - 1] : Math.max(0, 30 - (p - 5)); }
function norm(s) { return String(s == null ? '' : s).replace(/ /g, ' ').toLocaleLowerCase('pl').replace(/\s+/g, ' ').trim(); }
function key(s) { return norm(s).split(' ').filter(Boolean).sort().join(' '); }
function titleCase(s) { return norm(s).replace(/(^|[\s\-'])(\S)/g, function (m, a, b) { return a + b.toLocaleUpperCase('pl'); }); }
function timeVal(t) { var m = /^(?:(\d+):)?(\d+(?:[.,]\d+)?)$/.exec(String(t || '').trim()); return m ? (+(m[1] || 0)) * 60 + parseFloat(m[2].replace(',', '.')) : null; }

/* ---------- PDF ---------- */
function parsePdf(buf) {
  var lib = window.pdfjsLib;
  if (!lib) return Promise.reject('Nie wczytano czytnika PDF.');
  lib.GlobalWorkerOptions.workerSrc = 'lib/pdf.worker.min.js';
  return lib.getDocument({ data: new Uint8Array(buf), isEvalSupported: false }).promise.then(function (doc) {
    var jobs = [];
    for (var i = 1; i <= doc.numPages; i++) jobs.push(doc.getPage(i).then(function (p) { return p.getTextContent(); }));
    return Promise.all(jobs);
  }).then(function (pages) {
    var out = { event: '', round: null, date: '', results: [], fastest: [] }, hashX = null, nameMax = null;
    pages.forEach(function (tc) {
      var it = tc.items.filter(function (x) { return x.str && x.str.trim(); })
        .map(function (x) { return { s: x.str.trim(), x: x.transform[4], y: x.transform[5] }; });
      function at(s) { for (var i = 0; i < it.length; i++) if (it[i].s === s) return it[i]; return null; }
      function below(h) { var best = null; it.forEach(function (x) { if (x.y < h.y - 3 && x.y > h.y - 25 && Math.abs(x.x - h.x) < 45 && (!best || x.y > best.y)) best = x; }); return best; }
      var ev = at('Event'); if (ev && !out.event) { var v = below(ev); if (v) out.event = v.s; }
      var dt = at('Date'); if (dt && !out.date) { var d = below(dt); if (d) out.date = d.s; }
      ['Fastest men:', 'Fastest woman:', 'Fastest men', 'Fastest woman'].forEach(function (lbl) {
        var l = at(lbl); if (!l) return;
        var row = it.filter(function (x) { return Math.abs(x.y - l.y) < 2 && x.x > l.x + 5 && x !== l; }).sort(function (a, b) { return a.x - b.x; });
        var txt = row.map(function (x) { return x.s; }).join(' '), m = /^(.*?)\s*\(\s*([^)]*?)\s*\)/.exec(txt);
        var nm = m ? m[1].trim() : txt.trim(), tm = m ? m[2] : '';
        if (nm && !/^-+$/.test(nm) && timeVal(tm) != null) out.fastest.push({ name: nm, time: tm, t: timeVal(tm), who: /woman/i.test(lbl) ? 'K' : 'M' });
      });
      var h = at('#'), dr = at('Driver');
      if (h) hashX = h.x;
      var firstCol = it.filter(function (x) { return /^(Q|W|R)\d+$/.test(x.s); }).sort(function (a, b) { return a.x - b.x; })[0];
      if (firstCol) nameMax = firstCol.x - 8;
      if (hashX == null) return;
      var top = h ? h.y - 3 : 1e9;
      it.forEach(function (x) {
        if (x.y >= top || Math.abs(x.x - hashX) > 12 || !/^\d{1,3}$/.test(x.s)) return;
        var nm = it.filter(function (y) { return Math.abs(y.y - x.y) < 1.5 && y.x > x.x + 15 && (nameMax == null || y.x < nameMax) && !/^\d+([.,:]\d+)*$/.test(y.s); })
          .sort(function (a, b) { return a.x - b.x; }).map(function (y) { return y.s; }).join(' ').trim();
        if (nm) out.results.push({ pos: +x.s, name: nm });
      });
    });
    out.results.sort(function (a, b) { return a.pos - b.pos; });
    var m = /(\d+)\s*$/.exec(out.event); out.round = m ? +m[1] : null;
    out.bestlap = out.fastest.slice().sort(function (a, b) { return a.t - b.t; })[0] || null;
    return out;
  });
}

/* ---------- XLSX: otwieranie ---------- */
function colIdx(l) { var n = 0; for (var i = 0; i < l.length; i++) n = n * 26 + l.charCodeAt(i) - 64; return n - 1; }
function colName(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; }
function splitRef(r) { var m = /^([A-Z]+)(\d+)$/.exec(r); return m ? { c: colIdx(m[1]), r: +m[2] } : null; }
function kids(el, name) { var o = []; for (var n = el.firstChild; n; n = n.nextSibling) if (n.nodeType === 1 && n.localName === name) o.push(n); return o; }
function kid(el, name) { return kids(el, name)[0] || null; }
function parseXml(s) { return new DOMParser().parseFromString(s, 'application/xml'); }
function ser(doc) { var s = new XMLSerializer().serializeToString(doc); return /^<\?xml/.test(s) ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s; }

var IDX = { 51: 'FFCC00', 52: 'FF9900', 53: 'FF6600' };
function hsl(hex) { var r = parseInt(hex.substr(0, 2), 16) / 255, g = parseInt(hex.substr(2, 2), 16) / 255, b = parseInt(hex.substr(4, 2), 16) / 255, mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2, h = 0, s = 0, d = mx - mn;
  if (d) { s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? ((g - b) / d + (g < b ? 6 : 0)) : mx === g ? ((b - r) / d + 2) : ((r - g) / d + 4); h *= 60; } return { h: h, s: s, l: l }; }
function isOrange(hex) { if (!hex) return false; var c = hsl(hex); return c.h >= 15 && c.h <= 50 && c.s >= .55 && c.l >= .35 && c.l <= .75; }

function openXlsx(buf) {
  return JSZip.loadAsync(buf).then(function (zip) {
    var W = {};
    function txt(p) { var f = zip.file(p); return f ? f.async('string') : Promise.resolve(null); }
    return Promise.all([txt('xl/workbook.xml'), txt('xl/_rels/workbook.xml.rels'), txt('xl/sharedStrings.xml'), txt('xl/styles.xml'), txt('xl/theme/theme1.xml')]).then(function (a) {
      W.zip = zip; W.wb = parseXml(a[0]); W.rels = parseXml(a[1]); W.styles = a[3] ? parseXml(a[3]) : null;
      W.ss = []; if (a[2]) { var sd = parseXml(a[2]); [].forEach.call(sd.getElementsByTagNameNS(NS, 'si'), function (si) { W.ss.push([].map.call(si.getElementsByTagNameNS(NS, 't'), function (t) { return t.textContent; }).join('')); }); }
      W.theme = []; if (a[4]) { var cs = parseXml(a[4]).getElementsByTagNameNS('*', 'clrScheme')[0]; if (cs) { var k = [].filter.call(cs.childNodes, function (n) { return n.nodeType === 1; }).map(function (n) { var c = n.getElementsByTagNameNS('*', 'srgbClr')[0] || n.getElementsByTagNameNS('*', 'sysClr')[0]; return c ? (c.localName === 'srgbClr' ? c.getAttribute('val') : c.getAttribute('lastClr') || '') : ''; }); W.theme = [k[1], k[0], k[3], k[2]].concat(k.slice(4)); } }
      var rel = {}; [].forEach.call(W.rels.getElementsByTagName('Relationship'), function (r) { rel[r.getAttribute('Id')] = r.getAttribute('Target'); });
      W.sheets = [].map.call(W.wb.getElementsByTagNameNS(NS, 'sheet'), function (s) { var t = rel[s.getAttributeNS(RNS, 'id') || s.getAttribute('r:id')] || ''; t = t.replace(/^\//, ''); return { name: s.getAttribute('name'), path: t.indexOf('xl/') === 0 ? t : 'xl/' + t }; });
      return Promise.all(W.sheets.map(function (s) { return txt(s.path).then(function (x) { s.doc = x ? parseXml(x) : null; }); }));
    }).then(function () { W.sheets = W.sheets.filter(function (s) { return s.doc; }); W.sheets.forEach(function (s) { indexSheet(W, s); }); styleInfo(W); return W; });
  });
}
function indexSheet(W, S) {
  S.data = S.doc.getElementsByTagNameNS(NS, 'sheetData')[0]; S.rows = {}; S.cells = {}; S.masters = {};
  kids(S.data, 'row').forEach(function (r) { S.rows[+r.getAttribute('r')] = r; kids(r, 'c').forEach(function (c) { S.cells[c.getAttribute('r')] = c;
    var f = kid(c, 'f'); if (f && f.getAttribute('t') === 'shared' && f.textContent) S.masters[f.getAttribute('si')] = { ref: c.getAttribute('r'), f: f.textContent }; }); });
}
function styleInfo(W) {
  W.xfs = []; W.fills = []; W.orangeXf = {}; W.oFill = null; if (!W.styles) return;
  var fills = W.styles.getElementsByTagNameNS(NS, 'fills')[0], xfs = W.styles.getElementsByTagNameNS(NS, 'cellXfs')[0];
  W.fillsEl = fills; W.xfsEl = xfs;
  kids(fills, 'fill').forEach(function (fl, i) { var pf = kid(fl, 'patternFill'), fg = pf && kid(pf, 'fgColor'), hex = '';
    if (pf && pf.getAttribute('patternType') && pf.getAttribute('patternType') !== 'none' && fg) { if (fg.getAttribute('rgb')) hex = fg.getAttribute('rgb').slice(-6); else if (fg.getAttribute('theme') != null) hex = W.theme[+fg.getAttribute('theme')] || ''; else if (fg.getAttribute('indexed') != null) hex = IDX[+fg.getAttribute('indexed')] || ''; var t = +(fg.getAttribute('tint') || 0); if (hex && t) { var c = hsl(hex), l = t < 0 ? c.l * (1 + t) : c.l * (1 - t) + t; if (l > .75 || l < .35) hex = ''; } }
    W.fills.push(isOrange(hex)); if (isOrange(hex) && W.oFill == null && hex.toUpperCase() === 'FFC000') W.oFill = i; });
  kids(xfs, 'xf').forEach(function (x) { W.xfs.push(x); });
}
function xfOrange(W, s) { var x = W.xfs[s]; return !!(x && W.fills[+(x.getAttribute('fillId') || 0)]); }
function orangeStyle(W, base) {
  if (!W.styles) return base; if (W.orangeXf[base] != null) return W.orangeXf[base];
  if (W.oFill == null) { var fl = W.styles.createElementNS(NS, 'fill'), pf = W.styles.createElementNS(NS, 'patternFill'), fg = W.styles.createElementNS(NS, 'fgColor'), bg = W.styles.createElementNS(NS, 'bgColor');
    pf.setAttribute('patternType', 'solid'); fg.setAttribute('rgb', 'FFFFC000'); bg.setAttribute('indexed', '64'); pf.appendChild(fg); pf.appendChild(bg); fl.appendChild(pf); W.fillsEl.appendChild(fl);
    W.oFill = W.fills.length; W.fills.push(true); W.fillsEl.setAttribute('count', W.fills.length); }
  var src = W.xfs[base] || W.xfs[0], nx = src.cloneNode(true); nx.setAttribute('fillId', W.oFill); nx.setAttribute('applyFill', '1');
  W.xfsEl.appendChild(nx); W.xfs.push(nx); W.xfsEl.setAttribute('count', W.xfs.length); W.orangeXf[base] = W.xfs.length - 1; return W.xfs.length - 1;
}

/* ---------- komórki ---------- */
function val(W, c) { if (!c) return ''; var t = c.getAttribute('t'), v = kid(c, 'v');
  if (t === 's') return v ? (W.ss[+v.textContent] || '') : ''; if (t === 'inlineStr') { var is = kid(c, 'is'); return is ? [].map.call(is.getElementsByTagNameNS(NS, 't'), function (x) { return x.textContent; }).join('') : ''; }
  if (!v) return ''; if (t === 'str' || t === 'e') return v.textContent; if (t === 'b') return v.textContent === '1'; var n = +v.textContent; return isNaN(n) ? v.textContent : n; }
function num(v) { if (typeof v === 'number') return v; if (v === '' || v == null || typeof v === 'boolean') return null; var n = parseFloat(String(v).replace(',', '.')); return isNaN(n) || /^#/.test(String(v)) ? null : n; }
function getRow(S, r) { if (S.rows[r]) return S.rows[r]; var el = S.doc.createElementNS(NS, 'row'); el.setAttribute('r', r);
  var after = null; Object.keys(S.rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (n) { if (n > r && !after) after = S.rows[n]; });
  S.data.insertBefore(el, after); S.rows[r] = el; return el; }
function getCell(S, ref) { if (S.cells[ref]) return S.cells[ref]; var p = splitRef(ref), row = getRow(S, p.r), el = S.doc.createElementNS(NS, 'c'); el.setAttribute('r', ref);
  var after = null; kids(row, 'c').forEach(function (c) { if (!after && splitRef(c.getAttribute('r')).c > p.c) after = c; }); row.insertBefore(el, after); S.cells[ref] = el;
  if (row.hasAttribute('spans')) row.removeAttribute('spans'); return el; }
function clear(c) { ['v', 'is', 'f'].forEach(function (n) { kids(c, n).forEach(function (x) { c.removeChild(x); }); }); c.removeAttribute('t'); }
function setNum(S, ref, n, style) { var c = getCell(S, ref); clear(c); var v = S.doc.createElementNS(NS, 'v'); v.textContent = String(n); c.appendChild(v); if (style != null) c.setAttribute('s', style); return c; }
function setStr(S, ref, s, style) { var c = getCell(S, ref); clear(c); c.setAttribute('t', 'inlineStr'); var is = S.doc.createElementNS(NS, 'is'), t = S.doc.createElementNS(NS, 't'); t.setAttribute('xml:space', 'preserve'); t.textContent = s; is.appendChild(t); c.appendChild(is); if (style != null) c.setAttribute('s', style); return c; }
function setBlank(S, ref, style) { var c = getCell(S, ref); clear(c); if (style != null) c.setAttribute('s', style); return c; }
function styleOf(S, ref) { var c = S.cells[ref]; return c ? +(c.getAttribute('s') || 0) : 0; }

/* ---------- formuły (SUM, MIN, MAX, SMALL, LARGE, + - * /) ---------- */
function shiftF(f, dr, dc) { return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])/g, function (m, a, col, b, row) { return a + (a ? col : colName(colIdx(col) + dc)) + b + (b ? row : (+row + dr)); }); }
function formulaOf(S, c) { var f = kid(c, 'f'); if (!f) return null; if (f.textContent) return f.textContent; var m = S.masters[f.getAttribute('si')]; if (!m) return null; var a = splitRef(m.ref), b = splitRef(c.getAttribute('r')); return shiftF(m.f, b.r - a.r, b.c - a.c); }
function evalF(W, S, f) {
  var s = f.replace(/^=/, '').replace(/\s+/g, ''), i = 0, ERR = { e: '#NUM!' };
  function cv(ref) { var c = S.cells[ref.replace(/\$/g, '')]; var v = val(W, c); if (typeof v === 'string' && /^#/.test(v)) throw { e: v }; return v; }
  function rng(a, b) { var p = splitRef(a.replace(/\$/g, '')), q = splitRef(b.replace(/\$/g, '')), o = []; for (var r = Math.min(p.r, q.r); r <= Math.max(p.r, q.r); r++) for (var c = Math.min(p.c, q.c); c <= Math.max(p.c, q.c); c++) { var v = cv(colName(c) + r); if (typeof v === 'number') o.push(v); } return o; }
  function arg() { var m = /^(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)/.exec(s.slice(i)); if (m) { i += m[0].length; return rng(m[1], m[2]); } var x = expr(); return [x]; }
  function factor() {
    if (s[i] === '-') { i++; return -factor(); } if (s[i] === '+') { i++; return factor(); }
    if (s[i] === '(') { i++; var x = expr(); i++; return x; }
    var m = /^\d+(\.\d+)?/.exec(s.slice(i)); if (m) { i += m[0].length; return +m[0]; }
    m = /^([A-Z]+)\(/.exec(s.slice(i));
    if (m) { var fn = m[1]; i += m[0].length; var args = [arg()]; while (s[i] === ',') { i++; args.push(arg()); } i++;
      var a = args[0], k = args[1] ? args[1][0] : 1;
      if (fn === 'SUM') return args.reduce(function (t, x) { return t + x.reduce(function (p, q) { return p + q; }, 0); }, 0);
      if (fn === 'MIN') return a.length ? Math.min.apply(null, a) : 0; if (fn === 'MAX') return a.length ? Math.max.apply(null, a) : 0;
      if (fn === 'SMALL' || fn === 'LARGE') { var srt = a.slice().sort(function (p, q) { return fn === 'SMALL' ? p - q : q - p; }); if (k < 1 || k > srt.length) throw ERR; return srt[k - 1]; }
      throw { u: 1 }; }
    m = /^\$?[A-Z]{1,3}\$?\d+/.exec(s.slice(i)); if (m) { i += m[0].length; var v = cv(m[0]); return typeof v === 'number' ? v : (v === '' ? 0 : (function () { throw ERR; })()); }
    throw { u: 1 };
  }
  function term() { var x = factor(); while (s[i] === '*' || s[i] === '/') { var o = s[i++], y = factor(); x = o === '*' ? x * y : x / y; } return x; }
  function expr() { var x = term(); while (s[i] === '+' || s[i] === '-') { var o = s[i++], y = term(); x = o === '+' ? x + y : x - y; } return x; }
  try { var r = expr(); if (i < s.length) return { u: 1 }; return { v: Math.round(r * 1e9) / 1e9 }; } catch (e) { return e.e ? { e: e.e } : { u: 1 }; }
}
function recalcRows(W, S, rows) {
  for (var pass = 0; pass < 3; pass++) rows.forEach(function (r) { var row = S.rows[r]; if (!row) return; kids(row, 'c').forEach(function (c) {
    var f = formulaOf(S, c); if (f == null) return; var res = evalF(W, S, f);
    kids(c, 'v').forEach(function (x) { c.removeChild(x); }); kids(c, 'is').forEach(function (x) { c.removeChild(x); });
    if (res.u) { c.removeAttribute('t'); return; }
    var v = S.doc.createElementNS(NS, 'v'); if (res.e) { c.setAttribute('t', 'e'); v.textContent = res.e; } else { c.removeAttribute('t'); v.textContent = String(res.v); } c.appendChild(v); }); });
}

/* ---------- tabele ---------- */
function headCols(W, S, r, c0, c1) { var ci = { rounds: {} }, row = S.rows[r]; if (!row) return ci;
  kids(row, 'c').forEach(function (c) { var p = splitRef(c.getAttribute('r')); if (c0 != null && (p.c < c0 || p.c > c1)) return; var raw = val(W, c), l = norm(raw);
    if (/zawodnik|kierowc|uczestnik|^liga$/.test(l) && ci.name == null) ci.name = p.c;
    if (/^lp\.?$|^miejsce$/.test(l) && ci.place == null) ci.place = p.c;
    if (/^total$|^suma$/.test(l)) ci.total = p.c;
    if (/total\s*\(-?\s*\d+\)/.test(l)) { ci.tn = ci.tn || []; ci.tn.push(p.c); }
    if (typeof raw === 'number' || /^\d+$/.test(l)) ci.rounds[+raw] = p.c; });
  return ci; }
function lastCol(ci) { var m = Math.max(ci.name, ci.place == null ? 0 : ci.place, ci.total == null ? 0 : ci.total); (ci.tn || []).forEach(function (c) { m = Math.max(m, c); }); Object.keys(ci.rounds).forEach(function (k) { m = Math.max(m, ci.rounds[k]); }); return m; }
function firstCol(ci) { var m = Math.min(ci.name, ci.place == null ? 1e9 : ci.place); (ci.tn || []).forEach(function (c) { m = Math.min(m, c); }); return m; }
function baseStyle(W, S, col, rows) { var cnt = {}; rows.forEach(function (r) { var s = styleOf(S, colName(col) + r); if (!xfOrange(W, s)) cnt[s] = (cnt[s] || 0) + 1; });
  var best = null; Object.keys(cnt).forEach(function (k) { if (best == null || cnt[k] > cnt[best]) best = k; }); return best == null ? 0 : +best; }

/* ---------- generalka ---------- */
function findGeneral(W) {
  for (var k = 0; k < W.sheets.length; k++) { var S = W.sheets[k], rs = Object.keys(S.rows).map(Number).sort(function (a, b) { return a - b; });
    for (var j = 0; j < Math.min(rs.length, 15); j++) { var ci = headCols(W, S, rs[j]); if (ci.name != null && Object.keys(ci.rounds).length && (ci.total != null || ci.tn)) return { S: S, hr: rs[j], ci: ci }; } }
  return null; }
function applyGeneral(W, P, opt) {
  var T = findGeneral(W); if (!T) throw 'W pliku generalki nie znaleziono tabeli (Lp. / Liga / numery rund / Total).';
  var S = T.S, ci = T.ci, rc = ci.rounds[opt.round]; if (rc == null) throw 'W generalce nie ma kolumny rundy ' + opt.round + '.';
  var c0 = firstCol(ci), c1 = lastCol(ci), rows = [], byKey = {}, maxLp = 0, last = T.hr;
  Object.keys(S.rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) { if (r <= T.hr) return; var nm = val(W, S.cells[colName(ci.name) + r]);
    if (String(nm).trim()) { rows.push(r); byKey[key(nm)] = r; last = r; if (ci.place != null) { var lp = num(val(W, S.cells[colName(ci.place) + r])); if (lp > maxLp) maxLp = lp; } } });
  if (!rows.length) throw 'Generalka nie ma jeszcze żadnych zawodników.';
  var tpl = rows[rows.length - 1], bases = {}; Object.keys(ci.rounds).forEach(function (k) { bases[ci.rounds[k]] = baseStyle(W, S, ci.rounds[k], rows); });
  var mul = opt.x2 ? 2 : 1, bl = P.bestlap ? key(P.bestlap.name) : null, inPdf = {}, rep = [];
  P.results.forEach(function (x) { inPdf[key(x.name)] = x; });
  // nowi zawodnicy
  P.results.forEach(function (x) { var k = key(x.name); if (byKey[k]) return; var r = ++last; rows.push(r); byKey[k] = r;
    for (var c = c0; c <= c1; c++) { var ref = colName(c) + r, tref = colName(c) + tpl, tc = S.cells[tref], st = tc ? +(tc.getAttribute('s') || 0) : null, f = tc ? formulaOf(S, tc) : null;
      if (bases[c] != null) st = bases[c];
      if (c === ci.name) setStr(S, ref, titleCase(x.name), st);
      else if (c === ci.place) setNum(S, ref, ++maxLp, st);
      else if (f) { var cell = setBlank(S, ref, st), fe = S.doc.createElementNS(NS, 'f'); fe.textContent = shiftF(f, r - tpl, 0); cell.appendChild(fe); }
      else { var rn = null; Object.keys(ci.rounds).forEach(function (k2) { if (ci.rounds[k2] === c) rn = +k2; });
        if (rn != null && rn < opt.round) setNum(S, ref, 0, st); else setBlank(S, ref, st); } }
    x.isNew = true; });
  // punkty w rundzie
  rows.forEach(function (r) { var nm = val(W, S.cells[colName(ci.name) + r]), k = key(nm), x = inPdf[k], ref = colName(rc) + r, p = 0, isBl = false;
    if (x) { p = pts(x.pos); if (bl && k === bl) { p += 1; isBl = true; } }
    setNum(S, ref, p * mul, isBl ? orangeStyle(W, bases[rc]) : bases[rc]);
    if (x) rep.push({ pos: x.pos, name: String(nm).trim(), pts: p * mul, bestlap: isBl, isNew: !!x.isNew }); });
  if (bl && !inPdf[bl]) rep.bestlapMissing = P.bestlap.name;
  recalcRows(W, S, rows);
  rep.sort(function (a, b) { return a.pos - b.pos; }); rep.absent = rows.length - rep.length; rep.sheet = S.name;
  return rep;
}

/* ---------- dywizje ---------- */
function findDivs(W) {
  var out = [], seed = null;
  W.sheets.forEach(function (S) { var rs = Object.keys(S.rows).map(Number).sort(function (a, b) { return a - b; });
    rs.forEach(function (r) {
      var row = S.rows[r], cells = kids(row, 'c'), lp = null, total = null;
      cells.forEach(function (c) { var l = norm(val(W, c)); if (/^lp\.?$/.test(l) && lp == null) lp = splitRef(c.getAttribute('r')).c; if (/^total$/.test(l)) total = splitRef(c.getAttribute('r')).c; });
      if (lp != null && total != null) { var ci = headCols(W, S, r, lp, lp + 40); if (ci.name != null) out.push({ S: S, hr: r, ci: ci }); }
      if (!seed) { var dv = null, zn = null; cells.forEach(function (c) { var l = norm(val(W, c)); if (l === 'dywizja') dv = splitRef(c.getAttribute('r')).c; if (/^zawodnik/.test(l)) zn = splitRef(c.getAttribute('r')).c; }); if (dv != null && zn != null) seed = { S: S, hr: r, dv: dv, zn: zn }; }
    }); });
  out.forEach(function (B, i) {
    var S = B.S, ci = B.ci, rs = Object.keys(S.rows).map(Number).sort(function (a, b) { return a - b; }), next = out[i + 1] && out[i + 1].S === S ? out[i + 1].hr : 1e9;
    B.slots = rs.filter(function (r) { return r > B.hr && r < next && num(val(W, S.cells[colName(ci.place) + r])) != null; });
    // tytuł bloku nad nagłówkiem
    B.title = ''; for (var u = B.hr - 1; u >= B.hr - 6 && !B.title; u--) { var row = S.rows[u]; if (!row) continue; kids(row, 'c').forEach(function (c) { var p = splitRef(c.getAttribute('r')).c, v = val(W, c); if (!B.title && p >= firstCol(ci) && typeof v === 'string' && /dywizj|liga|grupa/i.test(v)) B.title = v.replace(/\s+/g, ' ').trim(); }); }
    var lm = /["„”“']?\s*([A-Z])\s*["„”“']?\s*$/.exec(B.title); B.letter = lm ? lm[1] : String.fromCharCode(65 + i);
    B.title = B.title.replace(/["„”“']/g, '').replace(/\s+/g, ' ').trim().toLowerCase().replace(/^./, function (c) { return c.toUpperCase(); }).replace(/ ([a-z])$/, function (m, c) { return ' ' + c.toUpperCase(); });
  });
  var members = {};
  if (seed) { var S2 = seed.S, cur = null; Object.keys(S2.rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) { if (r <= seed.hr) return;
    var d = String(val(W, S2.cells[colName(seed.dv) + r]) || '').trim(), n = String(val(W, S2.cells[colName(seed.zn) + r]) || '').trim();
    if (/^[A-Za-z]$/.test(d)) cur = d.toUpperCase(); if (n && cur) (members[cur] = members[cur] || []).push(n); }); }
  return { blocks: out, seed: members };
}
function applyDivs(W, P, opt) {
  var D = findDivs(W); if (!D.blocks.length) throw 'W pliku dywizji nie znaleziono tabel dywizji (Lp. / Liga / numery rund / Total).';
  var mul = opt.x2 ? 2 : 1, pos = {}, rep = { divs: [], unassigned: [] }, seen = {};
  P.results.forEach(function (x) { pos[key(x.name)] = x; });
  D.blocks.forEach(function (B) {
    var S = B.S, ci = B.ci, rc = ci.rounds[opt.round], names = [], filled = false;
    B.slots.forEach(function (r) { var n = String(val(W, S.cells[colName(ci.name) + r]) || '').trim(); if (n) names.push({ r: r, n: n }); });
    if (!names.length && D.seed[B.letter]) { var free = B.slots.slice(); D.seed[B.letter].forEach(function (n) { var r = free.shift(); if (r == null) { rep.overflow = (rep.overflow || []).concat(n); return; } setStr(S, colName(ci.name) + r, n, styleOf(S, colName(ci.name) + r)); names.push({ r: r, n: n }); filled = true; }); }
    var dr = { name: B.title || ('Dywizja ' + B.letter), letter: B.letter, rows: [], filled: filled, noCol: rc == null };
    if (rc != null) {
      var base = baseStyle(W, S, rc, B.slots), raced = names.filter(function (m) { return pos[key(m.n)]; }).sort(function (a, b) { return pos[key(a.n)].pos - pos[key(b.n)].pos; });
      names.forEach(function (m) { var x = pos[key(m.n)], p = 0; if (x) { seen[key(m.n)] = 1; p = opt.divMode === 'all' ? pts(x.pos) : pts(raced.indexOf(m) + 1); }
        setNum(S, colName(rc) + m.r, p * mul, base); dr.rows.push({ name: m.n, pos: x ? x.pos : null, dpos: x ? raced.indexOf(m) + 1 : null, pts: p * mul }); });
      recalcRows(W, S, B.slots);
    }
    dr.rows.sort(function (a, b) { return (a.dpos || 99) - (b.dpos || 99); }); rep.divs.push(dr);
  });
  P.results.forEach(function (x) { if (!seen[key(x.name)]) rep.unassigned.push(x.name); });
  return rep;
}

/* ---------- zapis ---------- */
function saveXlsx(W) {
  var cp = W.wb.getElementsByTagNameNS(NS, 'calcPr')[0];
  if (!cp) { cp = W.wb.createElementNS(NS, 'calcPr'); var after = W.wb.getElementsByTagNameNS(NS, 'sheets')[0]; after.parentNode.insertBefore(cp, after.nextSibling); var dn = W.wb.getElementsByTagNameNS(NS, 'definedNames')[0]; if (dn) dn.parentNode.insertBefore(cp, dn.nextSibling); }
  cp.setAttribute('fullCalcOnLoad', '1');
  W.zip.file('xl/workbook.xml', ser(W.wb));
  if (W.styles) W.zip.file('xl/styles.xml', ser(W.styles));
  W.sheets.forEach(function (s) { W.zip.file(s.path, ser(s.doc)); });
  // calcChain przebuduje Excel
  return Promise.resolve().then(function () {
    if (!W.zip.file('xl/calcChain.xml')) return;
    W.zip.remove('xl/calcChain.xml');
    [].slice.call(W.rels.getElementsByTagName('Relationship')).forEach(function (r) { if (/calcChain/.test(r.getAttribute('Target'))) r.parentNode.removeChild(r); });
    W.zip.file('xl/_rels/workbook.xml.rels', ser(W.rels));
    return W.zip.file('[Content_Types].xml').async('string').then(function (ct) { W.zip.file('[Content_Types].xml', ct.replace(/<Override[^>]*calcChain[^>]*\/>/g, '')); });
  }).then(function () { return W.zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } }); });
}

window.CXImport = { parsePdf: parsePdf, openXlsx: openXlsx, applyGeneral: applyGeneral, applyDivs: applyDivs, saveXlsx: saveXlsx, pts: pts, key: key };
})();
