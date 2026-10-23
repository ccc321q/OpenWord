/* ==========================================================================
   OpenWord — DOCX (OOXML) reader
   Reads word/document.xml + styles.xml + numbering.xml + rels and produces the
   normalised HTML the editor works with. Also used for documents exported from
   Google Docs (they are plain OOXML packages).
   ========================================================================== */
WW.docxImport = (function (U) {
  'use strict';

  var W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  function localOf(el) { return (el.localName || el.nodeName.split(':').pop() || '').toLowerCase(); }
  function isW(el, name) {
    if (!el || el.nodeType !== 1) return false;
    var ns = el.namespaceURI || '';
    if (ns && ns.indexOf('wordprocessingml') < 0 && ns.indexOf('schemas.openxmlformats.org') < 0) return false;
    return localOf(el) === name;
  }
  function kids(el, name) {
    var out = [];
    if (!el) return out;
    for (var i = 0; i < el.childNodes.length; i++) {
      var c = el.childNodes[i];
      if (c.nodeType === 1 && (!name || localOf(c) === name)) out.push(c);
    }
    return out;
  }
  function first(el, name) { var k = kids(el, name); return k.length ? k[0] : null; }
  function attr(el, name) {
    if (!el) return null;
    var v = el.getAttributeNS ? (el.getAttributeNS(W_NS, name) || el.getAttribute('w:' + name)) : null;
    if (v === null || v === '') v = el.getAttribute(name);
    return v === null || v === '' ? null : v;
  }
  /** Skip markup-only containers so their payload is still read exactly once. */
  function transparentWrappers(el) {
    var ln = localOf(el);
    return ['ins', 'sdt', 'sdtcontent', 'smarttag', 'customxml', 'prooferr', 'bdo', 'dir',
      'bookmarkstart', 'bookmarkend', 'commentrangestart', 'commentrangeend', 'contentpart'].indexOf(ln) >= 0;
  }

  /* --------------------------------------------------------------- rels -- */
  function parseRels(xmlText) {
    var map = {};
    if (!xmlText) return map;
    var doc = new DOMParser().parseFromString(xmlText, 'application/xml');
    kids(doc.documentElement).forEach(function (r) {
      var id = r.getAttribute('Id');
      var target = (r.getAttribute('Target') || '').replace(/^\/word\//, '').replace(/^\.\//, '');
      if (target.indexOf('/') !== 0 && target.indexOf('media/') !== 0 && target.indexOf('..') < 0 && target.indexOf('/') < 0) {
        // relative to word/
        target = 'word/' + target;
      } else if (target.charAt(0) === '/') target = target.slice(1);
      if (target.indexOf('word/') !== 0 && target.indexOf('docProps/') !== 0) target = 'word/' + target;
      map[id] = { target: target, mode: r.getAttribute('TargetMode') || 'Internal', type: r.getAttribute('Type') || '' };
    });
    return map;
  }

  function partPathFromRels(xmlText) {
    var out = {};
    if (!xmlText) return out;
    try {
      var doc = new DOMParser().parseFromString(xmlText, 'application/xml');
      kids(doc.documentElement).forEach(function (r) {
        var type = (r.getAttribute('Type') || '').toLowerCase();
        var target = (r.getAttribute('Target') || '').replace(/^\//, '');
        if (type.indexOf('/officeDocument') > 0) out.main = target;
      });
    } catch (e) { /* ignore */ }
    return out;
  }

  /* ------------------------------------------------------------- styles -- */
  function parseStyles(xmlText) {
    var styles = {};
    if (!xmlText) return styles;
    var doc;
    try { doc = new DOMParser().parseFromString(xmlText, 'application/xml'); } catch (e) { return styles; }
    if (!doc.documentElement || localOf(doc.documentElement) !== 'styles') return styles;
    kids(doc.documentElement, 'style').forEach(function (st) {
      var id = attr(st, 'styleId') || st.getAttribute('w:styleId');
      if (!id) return;
      var nameEl = first(st, 'name');
      var name = nameEl ? attr(nameEl, 'val') : id;
      var rec = {
        id: id,
        name: name || id,
        type: st.getAttributeNS ? (st.getAttributeNS(W_NS, 'type') || st.getAttribute('w:type') || 'paragraph') : 'paragraph',
        basedOn: (function () { var b = first(st, 'basedOn'); return b ? attr(b, 'val') : null; })(),
        pPr: propFromPr(first(st, 'pPr')),
        rPr: runPropsFromPr(first(st, 'rPr'))
      };
      styles[id] = rec;
    });
    return styles;
  }

  function styleChain(styles, id, seen) {
    seen = seen || [];
    var out = [];
    var cur = id && styles[id];
    var guard = 0;
    while (cur && guard++ < 12 && seen.indexOf(cur.id) < 0) {
      seen.push(cur.id);
      out.unshift(cur);
      cur = cur.basedOn ? styles[cur.basedOn] : null;
    }
    return out;
  }

  function resolveStyle(styles, id) {
    var merged = { pPr: {}, rPr: {}, name: '', headingLevel: 0 };
    if (!id) return merged;
    styleChain(styles, id).forEach(function (st) {
      Object.keys(st.pPr).forEach(function (k) { merged.pPr[k] = st.pPr[k]; });
      Object.keys(st.rPr).forEach(function (k) { if (st.rPr[k] !== undefined && st.rPr[k] !== null) merged.rPr[k] = st.rPr[k]; });
      merged.name = st.name;
      var m = /^heading\s*([1-9])$/i.exec(st.name || '');
      if (m) merged.headingLevel = parseInt(m[1], 10);
      var m2 = /^heading\s*([1-9])$/i.exec(st.id || '');
      if (m2) merged.headingLevel = Math.max(merged.headingLevel, parseInt(m2[1], 10));
    });
    return merged;
  }

  /** w:pPr -> plain object (twips kept as numbers). */
  function propFromPr(pr) {
    var o = {};
    if (!pr) return o;
    var jc = first(pr, 'jc'); if (jc) o.jc = attr(jc, 'val');
    var spacing = first(pr, 'spacing');
    if (spacing) {
      var b = attr(spacing, 'before'), a = attr(spacing, 'after'), l = attr(spacing, 'line'), rule = attr(spacing, 'lineRule');
      if (b) o.before = parseInt(b, 10);
      if (a) o.after = parseInt(a, 10);
      if (l) { o.line = parseInt(l, 10); o.lineRule = rule || 'auto'; }
    }
    var ind = first(pr, 'ind');
    if (ind) {
      ['left', 'right', 'firstLine', 'hanging', 'start', 'end'].forEach(function (k) {
        var v = attr(ind, k);
        if (v !== null && !isNaN(parseFloat(v))) o[k] = parseFloat(v);
      });
    }
    var numPr = first(pr, 'numPr');
    if (numPr) {
      var ilvl = first(numPr, 'ilvl'), numId = first(numPr, 'numId');
      o.numPr = {
        ilvl: ilvl ? parseInt(attr(ilvl, 'val') || '0', 10) : 0,
        numId: numId ? (attr(numId, 'val') || '0') : '0'
      };
    }
    var shd = first(pr, 'shd');
    if (shd) { var fill = attr(shd, 'fill'); if (fill && fill !== 'auto') o.shd = fill; }
    var outline = first(pr, 'outlineLvl');
    if (outline) o.outlineLvl = parseInt(attr(outline, 'val'), 10);
    var pStyle = first(pr, 'pStyle');
    if (pStyle) o.pStyle = attr(pStyle, 'val');
    var pBdr = first(pr, 'pBdr');
    if (pBdr) {
      var bottom = first(pBdr, 'bottom');
      if (bottom && (attr(bottom, 'val') || 'single') !== 'none') o.borderBottom = { val: attr(bottom, 'val') || 'single', sz: parseInt(attr(bottom, 'sz') || '6', 10), color: attr(bottom, 'color') || '000000' };
    }
    var keepNext = first(pr, 'keepNext'); if (keepNext && attr(keepNext, 'val') !== 'false' && attr(keepNext, 'val') !== '0') o.keepNext = true;
    var pageBreakBefore = first(pr, 'pageBreakBefore'); if (pageBreakBefore && attr(pageBreakBefore, 'val') !== 'false') o.pageBreakBefore = true;
    return o;
  }

  function runPropsFromPr(pr) {
    var o = {};
    if (!pr) return o;
    function onoff(name) {
      var el = first(pr, name);
      if (!el) return undefined;
      var v = attr(el, 'val');
      return !(v === 'false' || v === '0' || v === 'none');
    }
    var b = onoff('b'); if (b !== undefined) o.bold = b;
    var i = onoff('i'); if (i !== undefined) o.italic = i;
    var s = onoff('strike'); if (s !== undefined) o.strike = s;
    var ds = onoff('dstrike'); if (ds !== undefined) o.strike = o.strike || ds;
    var u = first(pr, 'u');
    if (u) { var uv = attr(u, 'val') || 'single'; o.underline = uv === 'none' ? false : uv; }
    var color = first(pr, 'color');
    if (color) { var cv = attr(color, 'val'); if (cv && cv !== 'auto') o.color = '#' + String(cv).toUpperCase(); }
    var hl = first(pr, 'highlight');
    if (hl) { var hv = attr(hl, 'val'); if (hv && hv !== 'none') o.highlight = hv; }
    var sz = first(pr, 'sz'); if (sz) o.size = parseInt(attr(sz, 'val') || '22', 10) / 2;
    var fonts = first(pr, 'rFonts');
    if (fonts) {
      var f = attr(fonts, 'ascii') || attr(fonts, 'hAnsi') || attr(fonts, 'cs');
      if (f) o.font = f;
    }
    var va = first(pr, 'vertAlign');
    if (va) { var vv = attr(va, 'val'); if (vv && vv !== 'baseline') o.vertAlign = vv; }
    var rStyle = first(pr, 'rStyle'); if (rStyle) o.rStyle = attr(rStyle, 'val');
    var vanish = onoff('vanish'); if (vanish) o.vanish = true;
    return o;
  }

  /* ---------------------------------------------------------- numbering -- */
  function parseNumbering(xmlText) {
    var abstract = {}, nums = {};
    if (!xmlText) return { abstract: abstract, nums: nums };
    var doc;
    try { doc = new DOMParser().parseFromString(xmlText, 'application/xml'); } catch (e) { return { abstract: abstract, nums: nums }; }
    if (!doc.documentElement) return { abstract: abstract, nums: nums };
    kids(doc.documentElement, 'abstractNum').forEach(function (an) {
      var id = attr(an, 'abstractNumId');
      var levels = {};
      kids(an, 'lvl').forEach(function (lvl) {
        var ilvl = parseInt(attr(lvl, 'ilvl') || '0', 10);
        var fmtEl = first(lvl, 'numFmt');
        var textEl = first(lvl, 'lvlText');
        var ppr = first(lvl, 'pPr');
        var ind = ppr ? first(ppr, 'ind') : null;
        var rprEl = first(lvl, 'rPr');
        levels[ilvl] = {
          numFmt: fmtEl ? (attr(fmtEl, 'val') || 'decimal') : 'decimal',
          lvlText: textEl ? (attr(textEl, 'val') || '') : '',
          left: ind && attr(ind, 'left') ? parseFloat(attr(ind, 'left')) : null,
          hanging: ind && attr(ind, 'hanging') ? parseFloat(attr(ind, 'hanging')) : null,
          start: (function () { var s = first(lvl, 'start'); return s ? parseInt(attr(s, 'val') || '1', 10) : 1; })(),
          font: rprEl ? (function () { var rf = first(rprEl, 'rFonts'); return rf ? attr(rf, 'ascii') : null; })() : null
        };
      });
      abstract[id] = levels;
    });
    kids(doc.documentElement, 'num').forEach(function (n) {
      var numId = attr(n, 'numId');
      var abs = first(n, 'abstractNumId');
      var overrides = {};
      kids(n, 'lvlOverride').forEach(function (ov) {
        var ilvl = parseInt(attr(ov, 'ilvl') || '0', 10);
        var so = first(ov, 'startOverride');
        overrides[ilvl] = { start: so ? parseInt(attr(so, 'val') || '1', 10) : null };
      });
      nums[numId] = { abstractId: abs ? attr(abs, 'val') : null, overrides: overrides };
    });
    return { abstract: abstract, nums: nums };
  }

  var FMT_TO_CSS = {
    bullet: 'disc', decimal: 'decimal', lowerLetter: 'decimal', lowerletter: 'decimal',
    upperLetter: 'decimal', upperletter: 'decimal', lowerRoman: 'decimal', lowerroman: 'decimal',
    upperRoman: 'decimal', upperroman: 'decimal', none: 'none'
  };

  /** Map a numbering level to {ordered, listStyleType, marker}. */
  function listTypeFor(level) {
    if (!level) return { ordered: false, listStyleType: 'disc' };
    var f = String(level.numFmt || 'decimal');
    var lower = f.toLowerCase();
    if (lower === 'bullet') {
      var t = (level.lvlText || '\u2022').trim();
      var marker = t.charAt(0) || '\u2022';
      if (/[\uF000-\uF0FF]/.test(marker)) marker = '\u2022';   // Symbol/Wingdings private use
      return { ordered: false, listStyleType: 'disc', marker: marker };
    }
    if (lower === 'none') return { ordered: false, orderedNone: true, listStyleType: 'none' };
    return { ordered: true, numFmt: f, listStyleType: 'decimal' };
  }

  /* ------------------------------------------------------------ images -- */
  function loadMedia(zip, rels) {
    var jobs = [];
    Object.keys(rels).forEach(function (id) {
      var rel = rels[id];
      if (!rel || rel.mode !== 'Internal') return;
      if (!/\/image$/.test(rel.type) && !/\.(png|jpe?g|gif|bmp|webp|svg|tiff?)$/i.test(rel.target)) return;
      var file = zip.file(rel.target);
      if (!file) return;
      jobs.push(file.async('base64').then(function (b64) {
        var ext = U.extensionOf(rel.target);
        var mime = U.mimeFromExt(ext);
        mediaCache[id] = 'data:' + mime + ';base64,' + b64.replace(/\r?\n/g, '');
      }).catch(function () { /* ignore unreadable media */ }));
    });
    return Promise.all(jobs);
  }
  var mediaCache = {};

  /* ------------------------------------------------------------- runs ---- */
  function styleKey(rp) {
    return [rp.font || '', rp.size || '', rp.bold ? 1 : 0, rp.italic ? 1 : 0, rp.underline || '',
      rp.strike ? 1 : 0, rp.color || '', rp.highlight || '', rp.vertAlign || '', rp.rStyle || ''].join('|');
  }

  function wrapInline(html, rp) {
    if (!html) return '';
    var style = '';
    if (rp.font) style += 'font-family:' + rp.font + ';';
    if (rp.size) style += 'font-size:' + (Math.round(rp.size * 100) / 100) + 'pt;';
    if (rp.color) style += 'color:' + rp.color + ';';
    if (rp.highlight) style += 'background-color:' + (U.HIGHLIGHT_COLORS[rp.highlight] || '#FFFF00') + ';';
    if (style) html = '<span style="' + style + '">' + html + '</span>';
    if (rp.bold) html = '<b>' + html + '</b>';
    if (rp.italic) html = '<i>' + html + '</i>';
    if (rp.underline) html = '<u>' + html + '</u>';
    if (rp.strike) html = '<s>' + html + '</s>';
    if (rp.vertAlign === 'superscript') html = '<sup>' + html + '</sup>';
    if (rp.vertAlign === 'subscript') html = '<sub>' + html + '</sub>';
    return html;
  }

  /** Convert a paragraph's children to HTML, merging identically formatted runs. */
  function runsToHtml(pEl, ctx) {
    var pieces = [];      // {type:'text'|'html', ...}
    walkRunContainer(pEl, {}, ctx, pieces);
    var out = '', buf = '', bufKey = null;
    function flush() {
      if (buf) { out += wrapInline(U.escapeHtml(buf), bufProps); buf = ''; }
      bufProps = null; bufKey = null;
    }
    var bufProps = null;
    pieces.forEach(function (piece) {
      if (piece.type === 'raw') { flush(); out += piece.html; return; }
      var k = styleKey(piece.rp);
      if (k !== bufKey) { flush(); bufKey = k; bufProps = piece.rp; }
      buf += piece.text;
    });
    flush();
    return out;
  }

  function walkRunContainer(el, props, ctx, pieces) {
    for (var i = 0; i < el.childNodes.length; i++) {
      var node = el.childNodes[i];
      if (node.nodeType === 3) { if (node.nodeValue) pieces.push({ type: 'text', text: node.nodeValue, rp: props }); continue; }
      if (node.nodeType !== 1) continue;
      var ln = localOf(node);
      if (ln === 'rpr' || ln === 'ppr' || ln === 'sectpr') continue;
      if (ln === 'prooferr' || ln === 'bookmarkstart' || ln === 'bookmarkend' ||
        ln === 'commentrangestart' || ln === 'commentrangeend' || ln === 'permstart' || ln === 'permend' ||
        ln === 'lastrenderedpagebreak' || ln === 'commentreference') continue;

      if (ln === 'r') {
        var rprNode = first(node, 'rPr');
        var rp = Object.assign({}, props, runPropsFromPr(rprNode));
        if (rp.vanish) continue;
        var hasContent = false;
        for (var j = 0; j < node.childNodes.length; j++) {
          var rc = node.childNodes[j];
          if (rc.nodeType !== 1) continue;
          var rln = localOf(rc);
          if (rln === 't') {
            hasContent = true;
            pieces.push({ type: 'text', text: rc.textContent, rp: rp });
          } else if (rln === 'tab') {
            pieces.push({ type: 'text', text: '\t', rp: rp });
          } else if (rln === 'br' || rln === 'cr') {
            var bt = attr(rc, 'type');
            if (bt === 'page') pieces.push({ type: 'raw', html: '<div class="ww-pagebreak"></div>' });
            else pieces.push({ type: 'text', text: '\n', rp: rp });
          } else if (rln === 'nobreakhyphen') {
            pieces.push({ type: 'text', text: '\u2011', rp: rp });
          } else if (rln === 'drawing' || rln === 'pict') {
            var html = drawingToHtml(rc, ctx);
            if (html) pieces.push({ type: 'raw', html: html });
          } else if (rln === 'sym') {
            var ch = attr(rc, 'char');
            if (ch) { pieces.push({ type: 'text', text: symToChar(ch, rc), rp: rp }); }
          } else if (rln === 'object' || rln === 'OLEObject') {
            ctx.warnings.push('An embedded object could not be converted and was skipped.');
          }
        }
        if (!hasContent) { /* empty run */ }
        continue;
      }
      if (ln === 'hyperlink') {
        var rid = relIdOf(node);
        var href = rid && ctx.rels[rid] ? ctx.rels[rid].target : '#';
        var anchor = node.getAttribute('w:anchor') || attr(node, 'anchor');
        if (anchor && !href.match(/^#/) && href === '#') href = '#' + anchor;
        var inner = [];
        walkRunContainer(node, Object.assign({}, props, { rStyle: 'Hyperlink' }), ctx, inner);
        var htmlInner = piecesToHtml(inner);
        pieces.push({
          type: 'raw',
          html: '<a href="' + U.escapeHtml(href) + '">' + (htmlInner || U.escapeHtml(href)) + '</a>'
        });
        continue;
      }
      if (ln === 'fldsimple') {
        var instr = node.getAttribute('w:instr') || attr(node, 'instr') || '';
        pieces.push({ type: 'raw', html: fieldHtml(instr, node.textContent) });
        continue;
      }
      if (ln === 'fldchar' || ln === 'instrtext') {
        var st = attr(node, 'fldCharType');
        if (st === 'begin') ctx.fieldActive = '';
        if (ln === 'instrtext' && ctx.fieldActive !== undefined && ctx.fieldActive !== null) {
          ctx.fieldActive += node.textContent;
        }
        if (st === 'end') {
          var instrTxt = ctx.fieldActive || '';
          ctx.fieldActive = null;
          var fh = fieldHtml(instrTxt, '');
          if (fh) pieces.push({ type: 'raw', html: fh });
        }
        continue;
      }
      if (ln === 'ins') { walkRunContainer(node, props, ctx, pieces); continue; }
      if (ln === 'del') { continue; }                       // drop tracked deletions
      if (ln === 'sdt' || ln === 'sdtcontent' || ln === 'smarttag' || ln === 'customxml') { walkRunContainer(node, props, ctx, pieces); continue; }
      if (ln === 'fallback') continue;                       // mc:Fallback duplicates
      if (ln === 'alternatecontent' || ln === 'choice') { walkRunContainer(node, props, ctx, pieces); continue; }
      if (ln === 'txbxcontent') {
        // Text box: pull the paragraphs out as normal blocks.
        var blocks = [];
        kids(node).forEach(function (b) {
          if (localOf(b) === 'p') blocks.push(b);
          else if (localOf(b) === 'tbl') blocks.push(b);
        });
        ctx.extraBlocks = ctx.extraBlocks || [];
        blocks.forEach(function (b) { ctx.extraBlocks.push(b); });
        continue;
      }
      if (transparentWrappers(node)) { walkRunContainer(node, props, ctx, pieces); continue; }
      // Unknown container: walk it so text is never lost.
      walkRunContainer(node, props, ctx, pieces);
    }
  }

  function piecesToHtml(pieces) {
    var out = '', buf = '', bufProps = null, bufKey = null;
    function flush() { if (buf) { out += wrapInline(U.escapeHtml(buf), bufProps || {}); buf = ''; } bufProps = null; bufKey = null; }
    pieces.forEach(function (piece) {
      if (piece.type === 'raw') { flush(); out += piece.html; return; }
      var k = styleKey(piece.rp);
      if (k !== bufKey) { flush(); bufKey = k; bufProps = piece.rp; }
      buf += piece.text;
    });
    flush();
    return out;
  }

  function fieldHtml(instr, cached) {
    var s = String(instr || '').toUpperCase();
    var map = [['PAGE', 'page'], ['NUMPAGES', 'numpages'], ['SECTIONPAGES', 'numpages'], ['DATE', 'date'], ['TIME', 'time'], ['FILENAME', 'filename'], ['AUTHOR', 'author']];
    for (var i = 0; i < map.length; i++) {
      if (s.indexOf(map[i][0]) === 0) {
        var val = (cached || '').trim();
        if (map[i][1] === 'page' || map[i][1] === 'numpages') val = val || '1';
        return '<span class="ww-field" data-ww-field="' + map[i][1] + '" contenteditable="false">' + U.escapeHtml(val) + '</span>';
      }
    }
    return '';
  }

  function symToChar(hex, node) {
    try {
      var n = parseInt(hex, 16);
      if (isNaN(n)) return '';
      return String.fromCodePoint(n);
    } catch (e) { return ''; }
  }

  function relIdOf(el) {
    var v = el.getAttributeNS ? (el.getAttributeNS(NS_R_URI, 'id') || el.getAttribute('r:id')) : null;
    if (!v) v = el.getAttribute('id');
    return v || null;
  }
  var NS_R_URI = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  function drawingToHtml(node, ctx) {
    var blips = findAllLocal(node, 'blip');
    var i, rid = null;
    for (i = 0; i < blips.length && !rid; i++) rid = relIdOf(blips[i]);
    if (!rid) {
      var imagedata = findAllLocal(node, 'imagedata');
      for (i = 0; i < imagedata.length && !rid; i++) rid = relIdOf(imagedata[i]);
    }
    if (!rid || !mediaCache[rid]) {
      if (rid && !mediaCache[rid]) ctx.warnings.push('An image referenced by ' + rid + ' is missing from the package.');
      return '';
    }
    // Extent (EMU) for display size
    var extents = findAllLocal(node, 'extent');
    var cx = extents.length ? parseFloat(extents[0].getAttribute('cx') || attr(extents[0], 'cx') || attr(extents[0], 'cx')) : 0;
    var cy = extents.length ? parseFloat(extents[0].getAttribute('cy') || attr(extents[0], 'cy') || attr(extents[0], 'cy')) : 0;
    var style = '';
    if (cx > 0 && cy > 0 && !isNaN(cx) && !isNaN(cy)) {
      var maxC = Math.max(cx, cy);
      var cap = ctx.contentEmu || 6206490;
      var scale = maxC > cap ? cap / maxC : 1;
      var pxW = Math.round(U.emuToPt(cx * scale) / 0.75);
      var pxH = Math.round(U.emuToPt(cy * scale) / 0.75);
      if (pxW > 0 && pxH > 0) style = 'width:' + pxW + 'px;height:' + pxH + 'px;';
    }
    var docPrs = findAllLocal(node, 'docPr');
    var alt = docPrs.length ? (docPrs[0].getAttribute('descr') || docPrs[0].getAttribute('title') || docPrs[0].getAttribute('name') || '') : '';
    var wrap = findAllLocal(node, 'anchor').length ? 'float:left;margin:0 12px 12px 0;' : '';
    ctx.imageSeq = (ctx.imageSeq || 0) + 1;
    return '<img src="' + mediaCache[rid] + '" alt="' + U.escapeHtml(alt) + '" data-ww-seq="' + ctx.imageSeq +
      '" style="' + style + wrap + '">';
  }

  function findAllLocal(root, name) {
    var out = [];
    (function walk(el) {
      for (var i = 0; i < el.childNodes.length; i++) {
        var c = el.childNodes[i];
        if (c.nodeType !== 1) continue;
        if (localOf(c) === name) out.push(c);
        walk(c);
      }
    })(root);
    return out;
  }

  /* ------------------------------------------------------- paragraphs --- */
  function paragraphToBlock(pEl, ctx) {
    var pr = first(pEl, 'pPr');
    var pProps = propFromPr(pr);
    var styleId = pProps.pStyle;
    var resolved = resolveStyle(ctx.styles, styleId);
    var merged = mergeProps(resolved.pPr, pProps);
    var runDefaults = resolved.rPr || {};

    ctx.extraBlocks = [];
    var inner = runsToHtml(pEl, ctx);
    var extra = ctx.extraBlocks || [];
    ctx.extraBlocks = [];

    var numPr = merged.numPr;
    var el;
    var level = 0;
    if (numPr && numPr.numId && String(numPr.numId) !== '0') {
      el = document.createElement('p');
      el.className = 'ww-list-paragraph';
      var numInfo = ctx.numbering.nums[numPr.numId];
      var abs = numInfo ? ctx.numbering.abstract[numInfo.abstractId] : null;
      var lvl = abs ? abs[numPr.ilvl] : null;
      var lt = listTypeFor(lvl);
      el.dataset.wwList = lt.ordered ? 'ol' : 'ul';
      el.dataset.wwIlvl = String(numPr.ilvl || 0);
      el.dataset.wwNumId = String(numPr.numId);
      el.dataset.wwOrdered = lt.ordered ? '1' : '';
      if (!lt.ordered && !lt.orderedNone) el.dataset.wwMarker = lt.marker || '\u2022';
      if (lvl && lvl.left !== null) el.dataset.wwIndent = String(lvl.left);
      if (lvl && lvl.hanging !== null) el.dataset.wwHanging = String(lvl.hanging);
    } else {
      el = document.createElement(headingTag(merged, resolved));
    }

    applyParagraphStyle(el, merged, runDefaults, ctx, { isListItem: !!numPr });
    el.innerHTML = inner;
    if (!el.textContent.trim() && !el.querySelector('img,.ww-field')) {
      if (/^H[1-6]$/.test(el.tagName)) { /* keep empty headings */ }
    }
    var result = [el];
    extra.forEach(function (b) {
      if (localOf(b) === 'p') result = result.concat(paragraphToBlock(b, ctx));
      else if (localOf(b) === 'tbl') result.push(tableToHtml(b, ctx));
    });
    return result;
  }

  function headingTag(merged, resolved) {
    if (resolved.headingLevel) return 'h' + Math.min(6, resolved.headingLevel);
    if (merged.outlineLvl !== undefined && merged.outlineLvl >= 0 && merged.outlineLvl < 6) return 'h' + (merged.outlineLvl + 1);
    var name = (resolved.name || '').toLowerCase();
    if (name === 'title') return 'p';
    if (name === 'subtitle') return 'p';
    if (name === 'quote') return 'p';
    if (name === 'intense quote') return 'p';
    return 'p';
  }

  function semanticClassFor(resolved) {
    var name = (resolved.name || '').toLowerCase();
    if (name === 'title') return 'ww-title';
    if (name === 'subtitle') return 'ww-subtitle';
    if (name === 'quote') return 'ww-quote';
    if (name === 'intense quote') return 'ww-intense-quote';
    return null;
  }

  function mergeProps(base, extra) {
    var o = {};
    Object.keys(base || {}).forEach(function (k) { o[k] = base[k]; });
    Object.keys(extra || {}).forEach(function (k) { if (extra[k] !== undefined) o[k] = extra[k]; });
    return o;
  }

  function twipsToPx(tw) { return Math.round(tw / 15 * 100) / 100; }

  function applyParagraphStyle(el, props, runDefaults, ctx, opts) {
    opts = opts || {};
    var jc = props.jc;
    if (jc) {
      var map = { left: 'left', center: 'center', right: 'right', both: 'justify', start: 'left', end: 'right' };
      if (map[jc]) el.style.textAlign = map[jc];
    }
    if (props.before !== undefined) el.style.marginTop = twipsToPx(props.before) + 'px';
    if (props.after !== undefined) el.style.marginBottom = twipsToPx(props.after) + 'px';
    // Indentation of list items is carried by the <ul>/<ol> nesting instead.
    if (!opts.isListItem) {
      if (props.left !== undefined) el.style.marginLeft = twipsToPx(props.left) + 'px';
      if (props.firstLine !== undefined) el.style.textIndent = twipsToPx(props.firstLine) + 'px';
      if (props.hanging !== undefined) el.style.textIndent = (-twipsToPx(props.hanging)) + 'px';
    }
    if (props.right !== undefined) el.style.marginRight = twipsToPx(props.right) + 'px';
    if (props.line !== undefined) {
      if ((props.lineRule || 'auto') === 'auto') el.style.lineHeight = String(Math.round((props.line / 240) * 1000) / 1000);
      else el.style.lineHeight = twipsToPx(props.line) + 'px';
    }
    if (props.shd) el.style.backgroundColor = '#' + props.shd;
    if (props.borderBottom) el.style.borderBottom = '1px solid ' + colorCss(props.borderBottom.color);
    if (props.pageBreakBefore) el.classList.add('ww-pagebreak');
    var cls = null;
    if (props.pStyle) {
      var r = resolveStyle(ctx.styles, props.pStyle);
      cls = semanticClassFor(r);
      if (r.headingLevel) cls = null;
    }
    if (cls) el.classList.add(cls);
    // Preserve paragraph-level run defaults that Word applies through the style
    // (colour / size of headings) only when the runs do not define them.
    var styleAttr = '';
    if (runDefaults) {
      if (runDefaults.color && props.keepRunColor) styleAttr += 'color:' + runDefaults.color + ';';
      if (runDefaults.size && props.keepRunSize) styleAttr += 'font-size:' + runDefaults.size + 'pt;';
    }
    if (styleAttr) el.setAttribute('style', (el.getAttribute('style') || '') + styleAttr);
  }

  function colorCss(c) {
    if (!c) return '#000000';
    if (/^[0-9a-fA-F]{6}$/.test(c)) return '#' + c.toUpperCase();
    return c;
  }

  /* ------------------------------------------------------------ tables --- */
  function tableToHtml(tblEl, ctx) {
    var rows = kids(tblEl, 'tr');
    if (!rows.length) return null;
    var tblPr = first(tblEl, 'tblPr');
    var grid = first(tblEl, 'tblGrid');
    var gridCols = grid ? kids(grid, 'gridCol').map(function (g) { return parseFloat(attr(g, 'w') || '0'); }) : [];
    var totalTwips = gridCols.reduce(function (a, b) { return a + b; }, 0);
    var contentPx = ctx.contentWidthPx || 624;
    var scale = totalTwips ? contentPx / totalTwips : 1;
    var tblW = tblPr ? first(tblPr, 'tblW') : null;
    var tblWVal = tblW ? parseFloat(attr(tblW, 'w') || '0') : 0;
    var tblWType = tblW ? (attr(tblW, 'type') || 'auto') : 'auto';
    var tableWidthPx = tblWVal ? (tblWType === 'pct' ? Math.round(contentPx * tblWVal / 5000) : Math.round(tblWVal * scale))
      : Math.round(totalTwips * scale);
    if (!tableWidthPx) tableWidthPx = Math.min(contentPx, 624);

    var table = document.createElement('table');
    table.style.width = tableWidthPx + 'px';
    var tblBorders = tblPr ? first(tblPr, 'tblBorders') : null;
    var borderDesc = bordersFromXml(tblBorders);
    if (borderDesc === 'none') { table.setAttribute('data-borders', 'none'); table.style.borderStyle = 'none'; }

    var colgroup = document.createElement('colgroup');
    gridCols.forEach(function (w) {
      var col = document.createElement('col');
      col.setAttribute('width', Math.max(4, Math.round(w * scale)));
      colgroup.appendChild(col);
    });
    if (gridCols.length) table.appendChild(colgroup);

    // owners[col] = the <td> that started the vertical merge covering that column
    var owners = {};
    var headerRow = tblPr ? first(tblPr, 'tblLook') : null;
    var firstRowHeader = headerRow ? (attr(headerRow, 'firstRow') === '1' || attr(headerRow, 'firstRow') === 'true') : false;

    rows.forEach(function (trEl, ri) {
      var trHtml = document.createElement('tr');
      var cells = kids(trEl, 'tc');
      var col = 0;
      if (ri === 0 && firstRowHeader) trHtml.dataset.wwHeaderRow = '1';
      cells.forEach(function (tc) {
        var tcPr = first(tc, 'tcPr');
        var gridSpan = tcPr ? parseInt(attr(first(tcPr, 'gridSpan'), 'val') || '1', 10) : 1;
        if (!(gridSpan > 0)) gridSpan = 1;
        var vMergeEl = tcPr ? first(tcPr, 'vMerge') : null;
        var vMerge = vMergeEl ? (attr(vMergeEl, 'val') || 'continue') : null;
        var tcW = tcPr ? first(tcPr, 'tcW') : null;
        var wVal = tcW ? parseFloat(attr(tcW, 'w') || '0') : 0;
        var wType = tcW ? (attr(tcW, 'type') || 'dxa') : 'dxa';
        var widthPx = wVal ? (wType === 'pct' ? Math.round(tableWidthPx * wVal / 5000) : Math.round(wVal * scale)) : 0;

        // A continuation cell carries no content: the owner cell grows instead.
        if (vMerge === 'continue') {
          var owner = owners[col];
          if (owner) owner.rowSpan = (owner.rowSpan || 1) + 1;
          for (var g = 0; g < gridSpan; g++) { col++; }
          return;
        }

        var td = document.createElement('td');
        td.rowSpan = 1;
        if (gridSpan > 1) td.colSpan = gridSpan;
        if (widthPx) td.style.width = widthPx + 'px';
        if (vMerge === 'restart') {
          owners[col] = td;
          for (var g2 = 1; g2 < gridSpan; g2++) owners[col + g2] = td;
        } else {
          for (var g3 = 0; g3 < gridSpan; g3++) owners[col + g3] = null;
        }
        if (tcPr) {
          var shd = first(tcPr, 'shd');
          if (shd) { var fill = attr(shd, 'fill'); if (fill && fill !== 'auto' && fill.toLowerCase() !== 'auto') td.style.backgroundColor = colorCss(fill); }
          var va = first(tcPr, 'vAlign');
          if (va) { var vav = attr(va, 'val'); td.style.verticalAlign = vav === 'center' ? 'middle' : (vav || 'top'); }
          var tcb = first(tcPr, 'tcBorders');
          if (tcb) applyCellBorders(td, tcb, scale);
        }
        if (borderDesc && borderDesc !== 'none' && borderDesc.css) td.style.borderColor = borderDesc.css;

        var content = kids(tc).filter(function (c) { return localOf(c) === 'p' || localOf(c) === 'tbl'; });
        var appendedAny = false;
        content.forEach(function (c) {
          if (localOf(c) === 'p') {
            paragraphToBlock(c, ctx).forEach(function (blk) { td.appendChild(blk); appendedAny = true; });
          } else {
            var t = tableToHtml(c, ctx);
            if (t) { td.appendChild(t); appendedAny = true; }
          }
        });
        if (!appendedAny) td.appendChild(document.createElement('p'));
        trHtml.appendChild(td);
        col += gridSpan;
      });
      table.appendChild(trHtml);
    });

    normalizeTable(table, gridCols.length || null);
    return table;
  }

  /** Pad short rows so the table keeps a rectangular grid (defensive against bad files). */
  function normalizeTable(table, expectedCols) {
    var rows = table.rows;
    if (!rows.length) return;
    var maxCols = 0;
    for (var i = 0; i < rows.length; i++) {
      var n = 0;
      for (var j = 0; j < rows[i].cells.length; j++) n += rows[i].cells[j].colSpan || 1;
      maxCols = Math.max(maxCols, n);
    }
    var target = Math.max(maxCols, expectedCols || 0);
    if (!target) return;
    for (var r = 0; r < rows.length; r++) {
      var count = 0;
      for (var c = 0; c < rows[r].cells.length; c++) count += rows[r].cells[c].colSpan || 1;
      while (count < target) {
        var td = document.createElement('td');
        td.appendChild(document.createElement('p'));
        td.dataset.wwSynthetic = '1';
        rows[r].appendChild(td);
        count++;
      }
    }
  }

  function bordersFromXml(tblBorders) {
    if (!tblBorders) return null;
    var sides = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
    var allNone = true, firstColor = null;
    for (var i = 0; i < sides.length; i++) {
      var b = first(tblBorders, sides[i]);
      if (!b) continue;
      var val = (attr(b, 'val') || 'single').toLowerCase();
      if (val !== 'none' && val !== 'nil') {
        allNone = false;
        if (!firstColor) firstColor = colorCss(attr(b, 'color'));
      }
    }
    if (allNone) return 'none';
    return { css: firstColor || '#ababab' };
  }

  function applyCellBorders(td, tcb, scale) {
    ['top', 'left', 'bottom', 'right'].forEach(function (side) {
      var b = first(tcb, side);
      if (!b) return;
      var val = (attr(b, 'val') || 'single').toLowerCase();
      if (val === 'none' || val === 'nil') { td.style['border' + cap(side) + 'Style'] = 'none'; return; }
      var sz = parseFloat(attr(b, 'sz') || '4');          // eighths of a point
      td.style['border' + cap(side) + 'Style'] = val === 'double' ? 'double' : 'solid';
      td.style['border' + cap(side) + 'Width'] = Math.max(1, Math.round(sz / 8 * 0.75 * 4) / 4) + 'px';
      td.style['border' + cap(side) + 'Color'] = colorCss(attr(b, 'color'));
    });
  }
  function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

  /* --------------------------------------------------------- lists build - */
  /**
   * Word stores lists as flat paragraphs carrying numId/ilvl. Rebuild real
   * <ul>/<ol> nesting from that information.
   */
  function ListBuilder() {
    this.root = document.createDocumentFragment();
    this.stack = [];   // {el, list, ilvl, numId}
  }
  ListBuilder.prototype.currentParent = function () {
    return this.stack.length ? this.stack[this.stack.length - 1].li : this.root;
  };
  ListBuilder.prototype.closeTo = function (ilvl) {
    while (this.stack.length > ilvl) this.stack.pop();
  };
  ListBuilder.prototype.flush = function () { this.stack = []; };

  function appendListParagraph(builder, meta) {
    var ilvl = meta.ilvl;
    var wanted = meta.wanted;   // 'ul' | 'ol'
    var numId = meta.numId;
    // unwind deeper levels
    while (builder.stack.length > ilvl + 1) builder.stack.pop();
    // a different list type / instance at this level closes the previous one
    var guard = 0;
    while (builder.stack.length === ilvl + 1 &&
      (builder.stack[ilvl].tag !== wanted || String(builder.stack[ilvl].numId) !== String(numId)) && guard++ < 12) {
      builder.stack.pop();
    }
    while (builder.stack.length < ilvl + 1) {
      var depth = builder.stack.length;
      var listEl = document.createElement(wanted);
      listEl.dataset.wwNumId = String(numId);
      listEl.dataset.wwIlvl = String(depth);
      if (meta.listStyleType && wanted === 'ol') {
        listEl.dataset.wwListType = meta.listStyleType;
        listEl.style.listStyleType = meta.listStyleType;
      }
      if (wanted === 'ul' && meta.markerCss) listEl.style.listStyleType = meta.markerCss;
      // Indentation: use the numbering definition when present, else sensible defaults.
      var deltaTwips = null;
      if (meta.left !== null && meta.parentLeft !== null) deltaTwips = meta.left - meta.parentLeft;
      else if (meta.left !== null) deltaTwips = meta.left;
      if (deltaTwips !== null && deltaTwips >= 0) listEl.style.paddingLeft = Math.round(deltaTwips / 20 * 100) / 100 + 'pt';
      else if (depth > 0) listEl.style.paddingLeft = '18pt';
      var anchor = builder.stack.length ? builder.stack[builder.stack.length - 1].li : null;
      if (anchor) anchor.appendChild(listEl);
      builder.stack.push({ tag: wanted, numId: numId, ilvl: depth, list: listEl, li: null });
    }
    var frame = builder.stack[ilvl];
    var li = document.createElement('li');
    frame.li = li;
    frame.list.appendChild(li);
    return li;
  }

  /* ------------------------------------------------------------- main ---- */
  function parse(arrayBuffer, options) {
    options = options || {};
    var warnings = [];
    var zip;
    return JSZip.loadAsync(arrayBuffer).then(function (z) {
      zip = z;
      return zip;
    }).then(function () {
      return Promise.all([
        readText(zip, '_rels/.rels'),
        zip.file('word/_rels/document.xml.rels') ? readText(zip, 'word/_rels/document.xml.rels') : Promise.resolve('')
      ]).then(function (res) {
        var rootRels = parseRels(res[0]);
        var mainPart = null;
        Object.keys(rootRels).forEach(function (k) {
          if (/officeDocument$/.test(rootRels[k].type) && rootRels[k].mode !== 'External') mainPart = rootRels[k].target;
        });
        if (!mainPart || !zip.file(mainPart)) {
          if (zip.file('word/document.xml')) mainPart = 'word/document.xml';
          else if (zip.file('document.xml')) mainPart = 'document.xml';
          else throw new Error('This is not a Word document (no main document part found).');
        }
        return { mainPart: mainPart, docRels: parseRels(res[1]) };
      });
    }).then(function (info) {
      return Promise.all([
        readText(zip, info.mainPart),
        readText(zip, 'word/styles.xml'),
        readText(zip, 'word/numbering.xml'),
        readText(zip, 'docProps/core.xml'),
        readText(zip, 'docProps/app.xml'),
        loadMedia(zip, info.docRels)
      ]).then(function (parts) {
        return { info: info, documentXml: parts[0], stylesXml: parts[1], numberingXml: parts[2], coreXml: parts[3], appXml: parts[4] };
      });
    }).then(function (data) {
      var docXml = data.documentXml;
      if (!docXml) throw new Error('The document part could not be read.');
      var parsed = new DOMParser().parseFromString(docXml.replace(/<?xml[^>]*\?>/, ''), 'application/xml');
      // Re-parse with the namespace prefix intact: parse the whole thing instead.
      parsed = new DOMParser().parseFromString(docXml, 'application/xml');
      var root = parsed.documentElement;
      if (!root || localOf(root) !== 'document') {
        if (root && localOf(root) === 'html') throw new Error('This .docx file is damaged (main part is not OOXML).');
        throw new Error('Could not parse the document XML. It may be corrupt.');
      }
      var ctx = {
        styles: parseStyles(data.stylesXml),
        numbering: parseNumbering(data.numberingXml),
        rels: data.info.docRels,
        warnings: warnings,
        contentWidthPx: options.contentWidthPx || 624,
        contentEmu: (options.contentWidthPx || 624) * 9525
      };
      var body = first(root, 'body');
      if (!body) throw new Error('The document has no body.');

      var outHtml = buildBody(body, ctx);
      var sectPr = first(body, 'sectPr');
      var page = pageFromSect(sectPr, ctx);
      var hf = readHeadersFooters(sectPr, ctx);
      var meta = parseCoreProps(data.coreXml);
      return {
        html: outHtml,
        page: page,
        meta: meta,
        headerHtml: hf.header,
        footerHtml: hf.footer,
        warnings: warnings,
        imageCount: Object.keys(mediaCache).length
      };
    });
  }

  function buildBody(body, ctx) {
    var container = document.createElement('div');
    var builder = new ListBuilder();
    function flushList() {
      if (builder.stack.length) {
        // attach the outermost list
        var top = builder.stack[0].list;
        if (top.parentNode) top.parentNode.removeChild(top);
        container.appendChild(top);
      }
      builder.flush();
    }
    var items = kids(body);
    items.forEach(function (el) {
      var ln = localOf(el);
      if (ln === 'sectpr') return;
      if (ln === 'p') {
        var pr = propFromPr(first(el, 'pPr'));
        var numPr = pr.numPr;
        if (numPr && String(numPr.numId) !== '0') {
          var numInfo = ctx.numbering.nums[numPr.numId];
          var abs = numInfo ? ctx.numbering.abstract[numInfo.abstractId] : null;
          var ilvlNum = numPr.ilvl || 0;
          var lvl = abs ? abs[ilvlNum] : null;
          var parentLvl = ilvlNum > 0 && abs ? abs[ilvlNum - 1] : null;
          var lt = listTypeFor(lvl);
          var markerCss = null;
          if (!lt.ordered && lvl && lvl.lvlText) {
            var mk = String(lvl.lvlText).trim().charAt(0) || '';
            if (mk && /[\uF000-\uF0FF]/.test(mk)) mk = '\u2022';
            if (mk && mk !== '\u2022' && mk !== '\u25E6' && mk !== '\u25AA') markerCss = '"' + mk.replace(/"/g, '') + '"';
          }
          var li = appendListParagraph(builder, {
            ilvl: ilvlNum, numId: numPr.numId,
            wanted: lt.ordered ? 'ol' : 'ul',
            listStyleType: lt.ordered && lvl ? cssListTypeFromFmt(lvl.numFmt) : null,
            markerCss: markerCss,
            left: lvl && lvl.left !== null && lvl.left !== undefined ? lvl.left : null,
            parentLeft: parentLvl && parentLvl.left !== null && parentLvl.left !== undefined ? parentLvl.left : null
          });
          var blocks = paragraphToBlock(el, ctx);
          blocks.forEach(function (b) {
            if (b === blocks[0]) {
              var p = document.createElement('p');
              copyAttrs(b, p);
              p.innerHTML = b.innerHTML;
              li.appendChild(p);
            } else {
              li.appendChild(b);
            }
          });
          return;
        }
        flushList();
        paragraphToBlock(el, ctx).forEach(function (b) { container.appendChild(b); });
        return;
      }
      if (ln === 'tbl') {
        flushList();
        var t = tableToHtml(el, ctx);
        if (t) container.appendChild(t);
        return;
      }
      if (ln === 'sdt' || ln === 'sdtcontent' || ln === 'customxml') {
        var inner = kids(el);
        if (inner.length) {
          var sub = document.createElement('div');
          inner.forEach(function (c) {
            if (localOf(c) === 'p') paragraphToBlock(c, ctx).forEach(function (b) { sub.appendChild(b); });
          });
          Array.prototype.forEach.call(sub.childNodes, function (n) { container.appendChild(n.cloneNode(true)); });
        }
        return;
      }
      if (ln === 'altchunk' || ln === 'embeddedobject') {
        warnings.push('An embedded part (' + ln + ') is not supported and was skipped.');
      }
    });
    flushList();
    return container.innerHTML;
  }

  function cssListTypeFromFmt(fmt) {
    var map = {
      decimal: '', lowerLetter: 'lower-alpha', lowerletter: 'lower-alpha', upperLetter: 'upper-alpha',
      upperletter: 'upper-alpha', lowerRoman: 'lower-roman', lowerroman: 'lower-roman',
      upperRoman: 'upper-roman', upperroman: 'upper-roman', decimalZero: 'decimal-leading-zero',
      decimalleadingzero: 'decimal-leading-zero', none: 'none'
    };
    return map[String(fmt)] || map[String(fmt).toLowerCase()] || '';
  }

  function copyAttrs(from, to) {
    Array.prototype.forEach.call(from.attributes, function (a) {
      if (a.name === 'class' && /ww-list-paragraph/.test(a.value)) return;
      to.setAttribute(a.name, a.value);
    });
  }

  function readText(zip, path) {
    var f = zip.file(path);
    if (!f) {
      // tolerate different casing
      var found = null;
      zip.forEach(function (rel) { if (rel.toLowerCase() === path.toLowerCase()) found = rel; });
      if (!found) return Promise.resolve('');
      f = zip.file(found);
    }
    return f.async('string').catch(function () { return ''; });
  }

  function pageFromSect(sectPr, ctx) {
    var page = { size: 'Letter', orientation: 'portrait', margins: { top: 1, right: 1, bottom: 1, left: 1, header: 0.5, footer: 0.5 } };
    if (!sectPr) return page;
    var pgSz = first(sectPr, 'pgSz');
    if (pgSz) {
      var w = parseFloat(attr(pgSz, 'w') || '12240'), h = parseFloat(attr(pgSz, 'h') || '15840');
      var orient = attr(pgSz, 'orient');
      if (w > h || orient === 'landscape') { page.orientation = 'landscape'; var t = w; w = h; h = t; }
      var best = null, bestDiff = Infinity;
      Object.keys(WW.docxExport.PAGE_SIZES).forEach(function (name) {
        var s = WW.docxExport.PAGE_SIZES[name];
        var diff = Math.abs(s[0] - w) + Math.abs(s[1] - h);
        if (diff < bestDiff) { bestDiff = diff; best = name; }
      });
      page.size = bestDiff <= 200 ? best : 'Custom';
      page.widthTwips = w; page.heightTwips = h;
    }
    var pgMar = first(sectPr, 'pgMar');
    if (pgMar) {
      ['top', 'right', 'bottom', 'left', 'header', 'footer'].forEach(function (k) {
        var v = attr(pgMar, k);
        if (v !== null && !isNaN(parseFloat(v))) page.margins[k] = Math.round(parseFloat(v) / 1440 * 1000) / 1000;
      });
    }
    return page;
  }

  function readHeadersFooters(sectPr, ctx) {
    var result = { header: '', footer: '' };
    if (!sectPr) return result;
    kids(sectPr).forEach(function (ref) {
      var ln = localOf(ref);
      if (ln !== 'headerreference' && ln !== 'footerreference') return;
      var type = attr(ref, 'type') || 'default';
      if (type !== 'default') return;   // first-page/even headers are ignored for now
      var rid = relIdOf(ref);
      var rel = rid && ctx.rels[rid];
      if (!rel) return;
      var part = zipGlobal[rel.target];
      if (!part) return;
      try {
        var doc = new DOMParser().parseFromString(part, 'application/xml');
        var root = doc.documentElement;
        if (!root) return;
        var paras = kids(root, 'p');
        var container = document.createElement('div');
        paras.forEach(function (p) { paragraphToBlock(p, ctx).forEach(function (b) { container.appendChild(b); }); });
        var html = container.innerHTML;
        if (ln === 'headerreference') result.header = html;
        else result.footer = html;
      } catch (e) { /* ignore */ }
    });
    return result;
  }

  function parseCoreProps(xml) {
    var meta = { title: '', author: '', subject: '', keywords: '', comments: '' };
    if (!xml) return meta;
    try {
      var doc = new DOMParser().parseFromString(xml, 'application/xml');
      function t(name) {
        var list = doc.getElementsByTagName(name);
        return list.length ? (list[0].textContent || '').trim() : '';
      }
      meta.title = t('dc:title'); meta.subject = t('dc:subject'); meta.author = t('dc:creator');
      meta.keywords = t('cp:keywords'); meta.comments = t('dc:description');
    } catch (e) { /* ignore */ }
    return meta;
  }

  var zipGlobal = {};   // path -> text, filled by parseZipCache()

  /**
   * Because header/footer parts are read after the (async) zip walk, cache the
   * XML text of every XML part up front. Packages are small enough for this.
   */
  function cacheAllXmlParts(zip) {
    var jobs = [];
    zip.forEach(function (path, entry) {
      if (!/\.xml$/i.test(path) || entry.dir) return;
      jobs.push(entry.async('string').then(function (txt) { zipGlobal[path] = txt; }, function () { }));
    });
    return Promise.all(jobs);
  }

  /** Public entry point. */
  function parseFile(arrayBuffer, options) {
    mediaCache = {};
    zipGlobal = {};
    return JSZip.loadAsync(arrayBuffer).then(function (zip) {
      return cacheAllXmlParts(zip);
    }).then(function () {
      return parse(arrayBuffer, options);
    });
  }

  return {
    parseFile: parseFile,
    parse: parse,
    _internals: {
      parseStyles: parseStyles, parseNumbering: parseNumbering, propFromPr: propFromPr,
      runPropsFromPr: runPropsFromPr, fieldHtml: fieldHtml, listTypeFor: listTypeFor,
      normalizeTable: normalizeTable, cssListTypeFromFmt: cssListTypeFromFmt
    }
  };
})(WW.util);
