/* ==========================================================================
   OpenWord — OpenDocument (.odt) reader / writer
   ODT is the format Google Docs and LibreOffice use for their "OpenDocument"
   export, so it is supported both ways.  Written by hand on top of JSZip.
   ========================================================================== */
WW.odt = (function (U) {
  'use strict';

  var MIME = 'application/vnd.oasis.opendocument.text';
  var NS = {
    office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
    style: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
    text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
    draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
    fo: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
    svg: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
    xlink: 'http://www.w3.org/1999/xlink',
    dc: 'http://purl.org/dc/elements/1.1/',
    meta: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
    manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
    number: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0'
  };

  function ln(el) { return (el.localName || '').toLowerCase(); }
  function kids(el, name) {
    var out = [];
    if (!el) return out;
    for (var i = 0; i < el.childNodes.length; i++) {
      var c = el.childNodes[i];
      if (c.nodeType === 1 && (!name || ln(c) === name)) out.push(c);
    }
    return out;
  }
  function first(el, name) { var k = kids(el, name); return k.length ? k[0] : null; }
  function at(el, name) {
    if (!el) return null;
    var v = el.getAttribute(name);
    if (v === null) {
      for (var i = 0; i < el.attributes.length; i++) {
        var a = el.attributes[i];
        if (ln2(a.name) === name) { v = a.value; break; }
      }
    }
    return v;
  }
  function ln2(name) { var i = name.indexOf(':'); return (i >= 0 ? name.slice(i + 1) : name); }

  /* ============================== READ =============================== */
  function parseOdt(arrayBuffer) {
    var warnings = [], media = {}, zip;
    return JSZip.loadAsync(arrayBuffer).then(function (z) {
      zip = z;
      var f = zip.file('content.xml');
      if (!f) throw new Error('Not an OpenDocument file (content.xml missing).');
      var jobs = [f.async('string'),
      readIfExists(zip, 'styles.xml'), readIfExists(zip, 'meta.xml')];
      zip.forEach(function (path, entry) {
        if (entry.dir || !/^(Pictures|Thumbnails)\//i.test(path)) return;
        if (/^Thumbnails\//i.test(path)) return;
        jobs.push(entry.async('base64').then(function (b64) {
          media[path] = 'data:' + U.mimeFromExt(U.extensionOf(path)) + ';base64,' + b64.replace(/\r?\n/g, '');
        }));
      });
      return Promise.all(jobs);
    }).then(function (parts) {
      var contentXml = parts[0], stylesXml = parts[1], metaXml = parts[2];
      var contentDoc = new DOMParser().parseFromString(contentXml, 'application/xml');
      var stylesDoc = stylesXml ? new DOMParser().parseFromString(stylesXml, 'application/xml') : null;
      var styles = collectStyles(contentDoc, stylesDoc);
      var ctx = { styles: styles, media: media, warnings: warnings };
      var body = first(first(contentDoc.documentElement, 'body'), 'text');
      if (!body) throw new Error('The OpenDocument has no text body.');
      var host = document.createElement('div');
      kids(body).forEach(function (el) {
        var node = convertBlock(el, ctx);
        if (node) {
          if (Array.isArray(node)) node.forEach(function (n) { if (n) host.appendChild(n); });
          else host.appendChild(node);
        }
      });
      var hf = readMasterStyles(contentDoc, stylesDoc, ctx);
      var meta = { title: '', author: '', subject: '', keywords: '', comments: '' };
      if (metaXml) {
        try {
          var md = new DOMParser().parseFromString(metaXml, 'application/xml');
          var mroot = md.documentElement;
          meta.title = textOf(first(mroot, 'title'));
          meta.subject = textOf(first(mroot, 'subject'));
          meta.comments = textOf(first(mroot, 'description'));
          var kw = kids(mroot, 'keyword').map(textOf).join(', ');
          meta.keywords = kw;
          var initials = first(mroot, 'initial-creator');
          meta.author = textOf(initials);
        } catch (e) { /* ignore */ }
      }
      var page = readPageLayout(stylesDoc, contentDoc);
      return { html: host.innerHTML, page: page, meta: meta, headerHtml: hf.header, footerHtml: hf.footer, warnings: warnings };
    });
  }

  function textOf(el) { return el ? (el.textContent || '').trim() : ''; }

  function readIfExists(zip, path) {
    var f = zip.file(path);
    return f ? f.async('string').catch(function () { return ''; }) : Promise.resolve('');
  }

  /** Build a flat style map: name -> {pPr:{}, rPr:{}, family, parent, outlineLevel, listStyle} */
  function collectStyles(contentDoc, stylesDoc) {
    var styles = {};
    [contentDoc, stylesDoc].forEach(function (doc) {
      if (!doc || !doc.documentElement) return;
      ['automatic-styles', 'styles', 'master-styles'].forEach(function (partName) {
        var part = first(doc.documentElement, partName);
        if (!part) return;
        kids(part, 'style').forEach(function (st) { registerStyle(st, styles); });
        kids(part, 'list-style').forEach(function (ls) {
          var name = at(ls, 'name');
          var levels = {};
          kids(ls).forEach(function (lvl, idx) {
            var num = parseInt(at(lvl, 'level') || (idx + 1), 10);
            var lnn = ln(lvl);
            if (lnn === 'list-level-style-number') {
              var nf = at(lvl, 'num-format') || at(lvl, 'style:num-format') || '1';
              var suffix = at(lvl, 'insert-text') || at(first(lvl, 'list-level-properties'), 'insert-text') || '.';
              levels[num - 1] = { ordered: true, numFormat: nf, suffix: suffix };
            } else if (lnn === 'list-level-style-bullet') {
              var bullet = at(lvl, 'bullet-char') || at(first(lvl, 'list-level-properties'), 'bullet-char') || '\u2022';
              levels[num - 1] = { ordered: false, bullet: bullet };
            }
          });
          if (name) styles['LIST:' + name] = { listLevels: levels };
        });
      });
    });
    return styles;
  }

  function registerStyle(st, styles) {
    var name = at(st, 'name');
    if (!name) return;
    var rec = {
      family: at(st, 'family') || 'paragraph',
      parent: at(st, 'parent') || null,
      pPr: {}, rPr: {}, outlineLevel: 0, listStyle: null, displayName: name
    };
    var pp = first(st, 'paragraph-properties');
    if (pp) {
      ['margin-left', 'margin-right', 'margin-top', 'margin-bottom', 'text-align', 'line-height',
        'text-indent', 'background-color'].forEach(function (k) {
          var v = at(pp, k);
          if (v !== null) rec.pPr[k] = v;
        });
      var ml = at(pp, 'margin-left'); if (ml !== null) rec.pPr['fo:margin-left'] = ml;
    }
    var tp = first(st, 'text-properties');
    if (tp) {
      ['font-weight', 'font-style', 'color', 'font-size', 'font-family',
        'text-underline-style', 'text-line-through-style', 'text-position',
        'background-color', 'hyphenate'].forEach(function (k) {
          var v = at(tp, k);
          if (v !== null) rec.rPr[k] = v;
        });
    }
    var op = first(st, 'text-outline-level-properties') || first(st, 'outline-level');
    var ol = op ? at(op, 'level') : null;
    if (ol) rec.outlineLevel = parseInt(ol, 10) || 0;
    var ls = first(st, 'list-style');
    if (ls && at(ls, 'name')) rec.listStyle = at(ls, 'name');
    var dname = first(st, 'display-name') || at(st, 'display-name');
    if (dname && typeof dname === 'string') rec.displayName = dname;
    else if (dname && dname.textContent) rec.displayName = dname.textContent.trim();
    styles[name] = rec;
  }

  function styleChain(styles, name, seen) {
    seen = seen || {};
    var out = [];
    var cur = name && styles[name];
    var guard = 0;
    while (cur && !seen[cur.parent || ''] && guard++ < 10) {
      if (cur.parent) seen[cur.parent] = true;
      out.unshift(cur);
      cur = cur.parent ? styles[cur.parent] : null;
    }
    return out;
  }

  function resolvedParagraphStyle(styles, name) {
    var o = { pPr: {}, rPr: {}, outlineLevel: 0, listStyle: null, name: '' };
    if (!name) return o;
    styleChain(styles, name).forEach(function (st) {
      Object.keys(st.pPr).forEach(function (k) { o.pPr[k] = st.pPr[k]; });
      Object.keys(st.rPr).forEach(function (k) { o.rPr[k] = st.rPr[k]; });
      if (st.outlineLevel) o.outlineLevel = st.outlineLevel;
      if (st.listStyle) o.listStyle = st.listStyle;
      o.name = st.displayName || st.name || '';
    });
    return o;
  }

  function odtLengthToPx(v) {
    if (v === null || v === undefined) return null;
    var m = /^(-?[\d.]+)\s*(cm|mm|in|pt|pc|px|%)?$/i.exec(String(v).trim());
    if (!m) return null;
    var n = parseFloat(m[1]), unit = (m[2] || 'cm').toLowerCase();
    var pt = unit === 'cm' ? n / 2.54 * 72 : unit === 'mm' ? n / 25.4 * 72 : unit === 'in' ? n * 72 :
      unit === 'pt' ? n : unit === 'pc' ? n * 12 : n * 0.75;
    return Math.round(pt / 0.75 * 100) / 100;
  }

  function inlineStyleFromTextProps(rPr) {
    if (!rPr) return '';
    var s = '';
    if (rPr['font-weight'] && /^(bold|bolder|[6-9]00)$/i.test(String(rPr['font-weight']))) s += 'font-weight:bold;';
    else if (rPr['font-weight']) s += 'font-weight:' + rPr['font-weight'] + ';';
    if (rPr['font-style'] === 'italic' || rPr['font-style'] === 'oblique') s += 'font-style:italic;';
    if (rPr.color && rPr.color !== 'auto') s += 'color:' + rPr.color + ';';
    if (rPr['background-color'] && rPr['background-color'] !== 'auto') s += 'background-color:' + rPr['background-color'] + ';';
    if (rPr['font-size']) s += 'font-size:' + (/^\d+$/.test(rPr['font-size']) ? rPr['font-size'] + 'pt' : rPr['font-size']) + ';';
    if (rPr['font-family']) s += 'font-family:' + rPr['font-family'] + ';';
    var deco = [];
    if (rPr['text-underline-style'] && rPr['text-underline-style'] !== 'none') deco.push('underline');
    if (rPr['text-line-through-style'] && rPr['text-line-through-style'] !== 'none') deco.push('line-through');
    if (deco.length) s += 'text-decoration:' + deco.join(' ') + ';';
    if (rPr['text-position']) {
      var pos = String(rPr['text-position']);
      if (/^super/i.test(pos)) s += 'vertical-align:super;font-size:.72em;';
      else if (/^sub/i.test(pos)) s += 'vertical-align:sub;font-size:.72em;';
    }
    return s;
  }

  function blockStyleFromParagraphProps(pPr) {
    var s = '';
    if (!pPr) return s;
    function len(v) { return odtLengthToPx(v); }
    var ml = pPr['margin-left'] !== undefined ? pPr['margin-left'] : pPr['fo:margin-left'];
    var mr = pPr['margin-right'] !== undefined ? pPr['margin-right'] : pPr['fo:margin-right'];
    var mt = pPr['margin-top'] !== undefined ? pPr['margin-top'] : pPr['fo:margin-top'];
    var mb = pPr['margin-bottom'] !== undefined ? pPr['margin-bottom'] : pPr['fo:margin-bottom'];
    var l = len(ml); if (l) s += 'margin-left:' + l + 'px;';
    l = len(mr); if (l) s += 'margin-right:' + l + 'px;';
    l = len(mt); if (l !== null) s += 'margin-top:' + l + 'px;';
    l = len(mb); if (l !== null) s += 'margin-bottom:' + l + 'px;';
    var ti = pPr['text-indent']; l = len(ti); if (l) s += 'text-indent:' + l + 'px;';
    if (pPr['text-align']) {
      var ta = pPr['text-align'];
      if (ta === 'justify' || ta === 'start' || ta === 'end') ta = ta === 'start' ? 'left' : ta === 'end' ? 'right' : 'justify';
      s += 'text-align:' + ta + ';';
    }
    if (pPr['line-height']) {
      var lh = String(pPr['line-height']);
      if (/%$/.test(lh)) s += 'line-height:' + (parseFloat(lh) / 100) + ';';
      else { var lpx = len(lh); if (lpx) s += 'line-height:' + lpx + 'px;'; }
    }
    if (pPr['background-color'] && pPr['background-color'] !== 'auto') s += 'background-color:' + pPr['background-color'] + ';';
    return s;
  }

  var ODT_HEADING_CLASSES = {
    'Title': 'ww-title', 'Subtitle': 'ww-subtitle', 'Quotations': 'ww-quote',
    'Preformatted_20_Text': null
  };

  function convertBlock(el, ctx) {
    var name = ln(el);
    if (name === 'p' || name === 'h') {
      var styleName = at(el, 'style-name');
      var st = resolvedParagraphStyle(ctx.styles, styleName);
      var tag = 'p';
      if (name === 'h') tag = 'h' + U.clamp(parseInt(at(el, 'outline-level') || '1', 10) || 1, 1, 6);
      else if (st.outlineLevel) tag = 'h' + U.clamp(st.outlineLevel, 1, 6);
      var out = [];
      var block = document.createElement(tag);
      var css = blockStyleFromParagraphProps(st.pPr) + blockStyleFromParagraphProps(propsOfInline(el, ctx));
      if (css) block.setAttribute('style', css);
      var cls = st.name && ODT_HEADING_CLASSES[st.name.replace(/\s+/g, '_')] ? ODT_HEADING_CLASSES[st.name.replace(/\s+/g, '_')] : null;
      if (cls) block.className = cls;
      block.innerHTML = inlineInner(el, ctx);
      out.push(block);
      return out;
    }
    if (name === 'list') {
      var styleAttr = at(el, 'style-name');
      var levels = (styleAttr && ctx.styles['LIST:' + styleAttr] && ctx.styles['LIST:' + styleAttr].listLevels) || null;
      if (!levels && styleAttr && ctx.styles[styleAttr]) {
        var rs = resolvedParagraphStyle(ctx.styles, styleAttr);
        if (rs.listStyle && ctx.styles['LIST:' + rs.listStyle]) levels = ctx.styles['LIST:' + rs.listStyle].listLevels;
      }
      var items = kids(el, 'list-item');
      if (!items.length) return null;
      var info = levels ? (levels[0] || levels[1]) : null;
      var ordered = info ? !!info.ordered : false;
      var list = document.createElement(ordered ? 'ol' : 'ul');
      if (ordered) {
        var cs = cssFromNumFormat(info.numFormat);
        if (cs) { list.dataset.wwListType = cs; list.style.listStyleType = cs; }
      } else if (info && info.bullet && info.bullet !== '\u2022') {
        list.style.listStyleType = '"' + info.bullet + '"';
      }
      list.dataset.wwOdtLevel = '0';
      items.forEach(function (item) {
        var li = document.createElement('li');
        var childBlocks = kids(item).filter(function (c) {
          return ['p', 'h', 'list', 'table', 'h'].indexOf(ln(c)) >= 0;
        });
        if (!childBlocks.length) li.appendChild(document.createElement('p'));
        childBlocks.forEach(function (c) {
          if (ln(c) === 'list') {
            var sub = convertBlock(c, Object.assign({}, ctx, {}));
            if (sub) {
              var subEl = Array.isArray(sub) ? sub[0] : sub;
              // re-derive the ordered/unordered flag for the nested level
              li.appendChild(subEl);
            }
            return;
          }
          var res = convertBlock(c, ctx);
          if (Array.isArray(res)) res.forEach(function (n) { if (n) li.appendChild(n); });
          else if (res) li.appendChild(res);
        });
        if (!li.querySelector('p,h1,h2,h3,h4,h5,h6,ul,ol,table,div')) {
          var emptyP = document.createElement('p');
          emptyP.innerHTML = li.innerHTML;
          li.innerHTML = '';
          li.appendChild(emptyP);
        }
        list.appendChild(li);
      });
      return list;
    }
    if (name === 'table') {
      return convertTable(el, ctx);
    }
    if (name === 'section' || name === 'soft-page-break') {
      if (name === 'soft-page-break') return null;
      var host = document.createDocumentFragment();
      kids(el).forEach(function (c) {
        var r = convertBlock(c, ctx);
        if (Array.isArray(r)) r.forEach(function (n) { if (n) host.appendChild(n); });
        else if (r) host.appendChild(r);
      });
      var wrap = document.createElement('div');
      wrap.appendChild(host);
      return Array.prototype.slice.call(wrap.childNodes);
    }
    if (name === 'alphabetical-index' || name === 'table-of-content' || name === 'notes') {
      ctx.warnings.push('An index/table of contents was replaced by its text.');
      var p = document.createElement('p');
      p.textContent = el.textContent;
      return p;
    }
    if (name === 'p') return null;
    return null;
  }

  function propsOfInline() { return {}; }

  function cssFromNumFormat(nf) {
    nf = String(nf || '');
    if (nf === '1' || nf === '') return '';
    if (nf === 'a') return 'lower-alpha';
    if (nf === 'A') return 'upper-alpha';
    if (nf === 'i') return 'lower-roman';
    if (nf === 'I') return 'upper-roman';
    return '';
  }

  function convertTable(el, ctx) {
    var table = document.createElement('table');
    var cols = kids(el, 'table-column');
    var colCount = 0;
    cols.forEach(function (c) { colCount += parseInt(at(c, 'number-columns-repeated') || '1', 10); });
    var colgroup = document.createElement('colgroup');
    var colWidths = [];
    cols.forEach(function (c) {
      var rep = parseInt(at(c, 'number-columns-repeated') || '1', 10);
      var w = at(c, 'column-width') || at(c, 'default-column-width');
      var px = odtLengthToPx(w);
      for (var i = 0; i < rep && colWidths.length < 200; i++) colWidths.push(px);
    });
    colWidths.forEach(function (px) {
      var col = document.createElement('col');
      if (px) col.setAttribute('width', Math.round(px));
      colgroup.appendChild(col);
    });
    if (colgroup.childNodes.length) table.appendChild(colgroup);
    var rows = kids(el, 'table-row');
    rows.forEach(function (rEl) {
      var repeatRow = parseInt(at(rEl, 'number-rows-repeated') || '1', 10);
      if (repeatRow > 100) repeatRow = 100;
      for (var rr = 0; rr < repeatRow; rr++) {
        var tr = document.createElement('tr');
        kids(rEl, 'table-cell').forEach(function (cEl) {
          var cRepeat = parseInt(at(cEl, 'number-columns-repeated') || '1', 10);
          var rSpan = parseInt(at(cEl, 'number-rows-spanned') || '1', 10);
          if (cRepeat > 200) cRepeat = 200;
          var covered = at(cEl, 'covered-table-cell');
          var td = document.createElement('td');
          if (cRepeat > 1 && !covered) td.colSpan = cRepeat;
          if (rSpan > 1) td.rowSpan = rSpan;
          var cs = tableCellStyle(cEl, ctx);
          if (cs) td.setAttribute('style', cs);
          var blocks = kids(cEl).filter(function (b) { return ['p', 'h', 'list', 'table'].indexOf(ln(b)) >= 0; });
          if (!blocks.length) td.appendChild(document.createElement('p'));
          blocks.forEach(function (b) {
            var res = convertBlock(b, ctx);
            if (Array.isArray(res)) res.forEach(function (n) { if (n) td.appendChild(n); });
            else if (res) td.appendChild(res);
          });
          for (var k = 0; k < cRepeat; k++) tr.appendChild(k === 0 ? td : td.cloneNode(true));
        });
        if (tr.childNodes.length) table.appendChild(tr);
      }
    });
    return table;
  }

  function tableCellStyle(cEl, ctx) {
    var styleName = at(cEl, 'style-name');
    var s = '';
    var st = ctx.styles[styleName];
    if (st) {
      var pp = resolvedParagraphStyle(ctx.styles, styleName).pPr;
      var chain = styleChain(ctx.styles, styleName);
      chain.forEach(function (c) {
        var tp = c.rPr || {};
        s += inlineStyleFromTextProps(tp);
      });
      if (pp['background-color']) s += 'background-color:' + pp['background-color'] + ';';
    }
    var cb = first(cEl, 'cell-properties') || (st ? null : null);
    if (cb) {
      var bg = at(cb, 'background-color');
      if (bg && bg !== 'auto') s += 'background-color:' + bg + ';';
      var va = at(cb, 'vertical-align');
      if (va && va !== 'auto') s += 'vertical-align:' + (va === 'middle' ? 'middle' : va) + ';';
      ['border-left', 'border-right', 'border-top', 'border-bottom'].forEach(function (k) {
        var v = at(cb, k);
        if (v && v !== 'none') {
          var parts = v.split(/\s+/);
          s += k.replace('border-', 'border-') + ':' + (odtLengthToPx(parts[0]) || 1) + 'px ' + (parts[1] || 'solid') + ' ' + (parts[2] || '#ababab') + ';';
        } else if (v === 'none') {
          s += k + ':none;';
        }
      });
      var pad = at(cb, 'padding');
      if (pad) s += 'padding:' + (odtLengthToPx(pad) || 0) + 'px;';
    }
    return s;
  }

  /** Inline children of a block -> HTML */
  function inlineInner(el, ctx) {
    var out = '';
    for (var i = 0; i < el.childNodes.length; i++) {
      out += inlineNode(el.childNodes[i], ctx, {});
    }
    return out;
  }

  function inlineNode(node, ctx, props) {
    if (node.nodeType === 3) return U.escapeHtml(node.nodeValue);
    if (node.nodeType !== 1) return '';
    var name = ln(node);
    switch (name) {
      case 's': {
        var n = parseInt(at(node, 'c') || '1', 10) || 1;
        return U.escapeHtml(new Array(n + 1).join('\u00a0'));
      }
      case 'tab': return U.escapeHtml('\t');
      case 'line-break': return U.escapeHtml('\n');
      case 'span': {
        var sn = at(node, 'style-name');
        var css = sn ? inlineStyleFromTextProps(resolvedParagraphStyle(ctx.styles, sn).rPr) : '';
        var inner = inlineInner(node, ctx);
        return css ? '<span style="' + css + '">' + inner + '</span>' : inner;
      }
      case 'a': {
        var href = at(node, 'href') || '#';
        return '<a href="' + U.escapeHtml(href) + '">' + inlineInner(node, ctx) + '</a>';
      }
      case 'phrase': case 'ins': return inlineInner(node, ctx);
      case 'strikethrough': return '<s>' + inlineInner(node, ctx) + '</s>';
      case 'soft-page-break': return '';
      case 'bookmark-start': case 'bookmark-end': case 'reference-mark-start': case 'reference-mark-end': return '';
      case 'note': {
        ctx.warnings.push('Footnotes/endnotes are kept as end-of-document notes.');
        var noteBody = first(node, 'note-body');
        var kind = at(node, 'class') === 'endnote' ? 'Endnote' : 'Footnote';
        var txt = noteBody ? noteBody.textContent.replace(/\s+/g, ' ').trim() : '';
        return '<i>[note: ' + U.escapeHtml(txt) + ']</i>';
      }
      case 'page-number': return '<span class="ww-field" data-ww-field="page" contenteditable="false">' + U.escapeHtml(node.textContent || '1') + '</span>';
      case 'page-count': return '<span class="ww-field" data-ww-field="numpages" contenteditable="false">' + U.escapeHtml(node.textContent || '1') + '</span>';
      case 'date': return '<span class="ww-field" data-ww-field="date" contenteditable="false">' + U.escapeHtml(node.textContent || '') + '</span>';
      case 'time': return '<span class="ww-field" data-ww-field="time" contenteditable="false">' + U.escapeHtml(node.textContent || '') + '</span>';
      case 'object': case 'frame': {
        var img = findImage(node, ctx);
        return img || inlineInner(node, ctx);
      }
      case 'image': {
        var im = findImageFromEl(node, ctx);
        return im || '';
      }
      default:
        return inlineInner(node, ctx);
    }
  }

  function findImage(node, ctx) {
    var frames = kids(node, 'frame');
    var list = frames.length ? frames : [node];
    for (var i = 0; i < list.length; i++) {
      var image = first(list[i], 'image');
      if (image) {
        var w = at(list[i], 'width'), h = at(list[i], 'height');
        var style = '';
        var wp = odtLengthToPx(w), hp = odtLengthToPx(h);
        if (wp && hp) style = 'width:' + Math.round(wp) + 'px;height:' + Math.round(hp) + 'px;';
        return findImageFromEl(image, ctx, style);
      }
    }
    return '';
  }

  function findImageFromEl(image, ctx, style) {
    var href = at(image, 'href') || '';
    var data = ctx.media[href] || ctx.media[href.replace(/^\.\//, '')];
    if (!data) { ctx.warnings.push('Image "' + href + '" could not be read from the archive.'); return ''; }
    return '<img src="' + data + '" alt=""' + (style ? ' style="' + style + '"' : '') + '>';
  }

  function readMasterStyles(contentDoc, stylesDoc, ctx) {
    var res = { header: '', footer: '' };
    var doc = contentDoc;
    var masters = first(doc.documentElement, 'master-styles');
    if (!masters && stylesDoc) masters = first(stylesDoc.documentElement, 'master-styles');
    if (!masters) return res;
    var mp = first(masters, 'master-page');
    if (!mp) return res;
    ['header', 'footer'].forEach(function (key) {
      var el = first(mp, key);
      if (!el) return;
      var host = document.createElement('div');
      kids(el).forEach(function (b) {
        var r = convertBlock(b, ctx);
        if (Array.isArray(r)) r.forEach(function (n) { if (n) host.appendChild(n); });
        else if (r) host.appendChild(r);
      });
      res[key] = host.innerHTML;
    });
    return res;
  }

  function readPageLayout(stylesDoc, contentDoc) {
    var page = { size: 'Letter', orientation: 'portrait', margins: { top: 1, right: 1, bottom: 1, left: 1, header: 0.5, footer: 0.5 } };
    var pl = null;
    [contentDoc, stylesDoc].forEach(function (doc) {
      if (!doc || !doc.documentElement || pl) return;
      ['automatic-styles', 'styles'].forEach(function (partName) {
        var part = first(doc.documentElement, partName);
        if (!part || pl) return;
        kids(part, 'page-layout').forEach(function (p) {
          if (!pl) pl = first(p, 'page-layout-properties');
        });
      });
    });
    if (!pl) return page;
    var wpx = odtLengthToPx(at(pl, 'page-width')), hpx = odtLengthToPx(at(pl, 'page-height'));
    if (wpx && hpx) {
      var wIn = wpx / 96, hIn = hpx / 96;
      var orient = 'portrait';
      if (wIn > hIn) { orient = 'landscape'; var t = wIn; wIn = hIn; hIn = t; }
      page.orientation = orient;
      var best = null, bestDiff = Infinity;
      Object.keys(WW.docxExport.PAGE_SIZES).forEach(function (name) {
        var s = WW.docxExport.PAGE_SIZES[name];
        var diff = Math.abs(s[0] / 1440 - wIn) + Math.abs(s[1] / 1440 - hIn);
        if (diff < bestDiff) { bestDiff = diff; best = name; }
      });
      page.size = bestDiff < 0.08 ? best : 'Custom';
    }
    var margins = { top: 'margin-top', right: 'margin-right', bottom: 'margin-bottom', left: 'margin-left' };
    Object.keys(margins).forEach(function (k) {
      var px = odtLengthToPx(at(pl, margins[k]));
      if (px) page.margins[k] = Math.round(px / 96 * 1000) / 1000;
    });
    return page;
  }

  /* ============================== WRITE ============================== */
  function OdtWriter() {
    this.auto = [];          // {name, xml}
    this.autoMap = {};
    this.pSeq = 0; this.tSeq = 0; this.lSeq = 0;
    this.images = [];        // {path, bytes, mime}
    this.imageMap = {};
    this.warnings = [];
  }
  OdtWriter.prototype.styleFor = function (kind, props, render) {
    var key = kind + ':' + JSON.stringify(props);
    if (this.autoMap[key]) return this.autoMap[key];
    var name = (kind === 'p' ? 'P' : kind === 't' ? 'T' : 'L') + (++this.pSeq);
    this.autoMap[key] = name;
    this.auto.push({ name: name, xml: render(props, name) });
    return name;
  };

  function paraStyleXml(props, name) {
    var pp = '', tp = '';
    if (props.ml) pp += ' fo:margin-left="' + props.ml + '"';
    if (props.mr) pp += ' fo:margin-right="' + props.mr + '"';
    if (props.mt !== undefined) pp += ' fo:margin-top="' + props.mt + '"';
    if (props.mb !== undefined) pp += ' fo:margin-bottom="' + props.mb + '"';
    if (props.ti) pp += ' fo:text-indent="' + props.ti + '"';
    if (props.align) pp += ' fo:text-align="' + props.align + '"';
    if (props.line) pp += ' style:line-height="' + props.line + '" style:line-height-at-least="0cm"';
    if (props.bg) pp += ' fo:background-color="' + props.bg + '"';
    if (pp) pp = '<style:paragraph-properties' + pp + '/>';
    if (props.tcss) tp = '<style:text-properties' + props.tcss + '/>';
    return '<style:style style:name="' + name + '" style:family="paragraph">' + pp + tp + '</style:style>';
  }

  function textStyleXml(props, name) {
    var a = '';
    if (props.b) a += ' fo:font-weight="bold"';
    if (props.i) a += ' fo:font-style="italic"';
    if (props.color) a += ' fo:color="' + props.color + '"';
    if (props.bg) a += ' fo:background-color="' + props.bg + '"';
    if (props.sz) a += ' fo:font-size="' + props.sz + 'pt"';
    if (props.font) a += ' fo:font-family="' + U.escapeXml(props.font) + '"';
    if (props.u) a += ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"';
    if (props.s) a += ' style:text-line-through-style="solid" style:text-line-through-width="auto" style:text-line-through-color="font-color"';
    if (props.va === 'superscript') a += ' style:text-position="super 58%"';
    if (props.va === 'subscript') a += ' style:text-position="sub 58%"';
    return '<style:style style:name="' + name + '" style:family="text"><style:text-properties' + a + '/></style:style>';
  }

  function listStyleXml(props, name) {
    var lv = '';
    for (var i = 0; i < 9; i++) {
      if (props.ordered) {
        lv += '<text:list-level-style-number text:level="' + (i + 1) + '" text:style-name="Numbering_Symbols" style:num-format="' +
          numFormatChar(props.fmt) + '" text:insert-text="' + '.' + '">' +
          '<style:list-level-properties text:list-level-position-and-distance="' + ((i + 1) * 1.27) + 'cm"/>' +
          '</text:list-level-style-number>';
      } else {
        lv += '<text:list-level-style-bullet text:level="' + (i + 1) + '" text:style-name="Bullet_Symbols" text:bullet-char="' +
          U.escapeXml(bulletFor(i)) + '">' +
          '<style:list-level-properties text:list-level-position-and-distance="' + ((i + 1) * 1.27) + 'cm"/>' +
          '</text:list-level-style-bullet>';
      }
    }
    return '<text:list-style style:name="' + name + '">' + lv + '</text:list-style>';
  }

  function numFormatChar(fmt) {
    return ({ decimal: '1', lowerLetter: 'a', upperLetter: 'A', lowerRoman: 'i', upperRoman: 'I', decimalZero: '01', none: '' })[fmt] || '1';
  }
  function bulletFor(i) { return ['\u2022', '\u25E6', '\u25AA', '\u2022', '\u25E6', '\u25AA', '\u2022', '\u25E6', '\u25AA'][i % 9]; }

  function cmFromPx(px) { return (px / 96 * 2.54).toFixed(3) + 'cm'; }
  function cmFromPt(pt) { return (pt / 72 * 2.54).toFixed(4) + 'cm'; }

  function exportOdt(opts) {
    var w = new OdtWriter();
    var page = Object.assign({
      size: 'Letter', orientation: 'portrait',
      margins: { top: 1, right: 1, bottom: 1, left: 1, header: 0.5, footer: 0.5 }
    }, opts.page || {});
    var size = WW.docxExport.PAGE_SIZES[page.size] || WW.docxExport.PAGE_SIZES.Letter;
    var wIn = size[0] / 1440, hIn = size[1] / 1440;
    if (page.orientation === 'landscape') { var t = wIn; wIn = hIn; hIn = t; }

    var bodyXml = blocksToOdt(opts.editor, w);
    var headerXml = hfToOdt(opts.header, w);
    var footerXml = hfToOdt(opts.footer, w);

    var listStyles = w.listStyles || {};

    var content = '<?xml version="1.0" encoding="UTF-8"?>' +
      '<office:document-content ' + nsDecl() + '>' +
      '<office:automatic-styles>' +
      '<style:style style:name="Standard" style:family="paragraph" style:class="text"/>' +
      '<style:style style:name="Text_20_body" style:display-name="Text Body" style:family="paragraph" style:parent-style-name="Standard" style:class="text"/>' +
      '<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body" style:class="text"><style:outline-level-properties/></style:style>' +
      w.auto.map(function (a) { return a.xml; }).join('') +
      Object.keys(listStyles).map(function (k) { return listStyles[k]; }).join('') +
      '<style:style style:name="Page_20_number" style:family="text"><style:text-properties style:font-name="Calibri" fo:size="10pt"/></style:style>' +
      '</office:automatic-styles>' +
      '<office:body><office:text text:use-soft-page-breaks="true">' + bodyXml + '</office:text></office:body>' +
      '<office:master-styles>' +
      '<style:master-page style:name="Standard" style:page-layout-name="pm1">' +
      (headerXml ? '<style:header>' + headerXml + '</style:header>' : '') +
      (footerXml ? '<style:footer>' + footerXml + '</style:footer>' : '') +
      '</style:master-page></office:master-styles>' +
      '</office:document-content>';

    var styles = '<?xml version="1.0" encoding="UTF-8"?>' +
      '<office:document-styles ' + nsDecl() + '>' +
      '<office:fonts/>' +
      '<office:styles>' +
      '<style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-top="0cm" fo:margin-bottom="0.247cm"/><style:text-properties fo:font-size="11pt" style:font-name="Calibri"/></style:default-style>' +
      '<style:style style:name="Normal" style:family="paragraph" style:class="text"/>' +
      '<style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="1" style:class="text">' +
      '<style:paragraph-properties fo:margin-top="0.423cm" fo:margin-bottom="0.212cm"/><style:text-properties fo:font-size="20pt" fo:color="#2b579a"/></style:style>' +
      '<style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="2" style:class="text">' +
      '<style:paragraph-properties fo:margin-top="0.388cm" fo:margin-bottom="0.176cm"/><style:text-properties fo:font-size="16pt" fo:color="#2b579a"/></style:style>' +
      '<style:style style:name="Heading_20_3" style:display-name="Heading 3" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="3" style:class="text">' +
      '<style:paragraph-properties fo:margin-top="0.353cm" fo:margin-bottom="0.141cm"/><style:text-properties fo:font-size="13.5pt" fo:color="#3b5b8c"/></style:style>' +
      '<style:style style:name="Quotations" style:display-name="Quotations" style:family="paragraph" style:parent-style-name="Standard" style:class="text">' +
      '<style:paragraph-properties fo:margin-left="1cm" fo:margin-right="1cm"/><style:text-properties fo:font-style="italic"/></style:style>' +
      '<style:style style:name="Title" style:family="paragraph" style:parent-style-name="Standard" style:class="text"><style:text-properties fo:font-size="26pt" fo:color="#2b579a"/></style:style>' +
      '<style:style style:name="Subtitle" style:family="paragraph" style:parent-style-name="Standard" style:class="text"><style:text-properties fo:font-size="13pt" fo:font-style="italic" fo:color="#5b7a9d"/></style:style>' +
      '<style:style style:name="List_20 Paragraph" style:display-name="List Paragraph" style:family="paragraph" style:parent-style-name="Text_20_body" style:class="list"/>' +
      '<style:style style:name="Numbering_Symbols" style:display-name="Numbering Symbols" style:family="text"/>' +
      '<style:style style:name="Bullet_Symbols" style:display-name="Bullet Symbols" style:family="text"><style:text-properties fo:font-family="OpenSymbol"/></style:style>' +
      '<style:style style:name="Internet_20_link" style:display-name="Internet link" style:family="text">' +
      '<style:text-properties fo:color="#0563c1" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>' +
      '</office:styles>' +
      '<office:automatic-styles>' +
      '<style:page-layout style:name="pm1"><style:page-layout-properties ' +
      'fo:page-width="' + cmFromIn(wIn) + '" fo:page-height="' + cmFromIn(hIn) + '" ' +
      'style:print-orientation="' + (page.orientation === 'landscape' ? 'landscape' : 'portrait') + '" ' +
      'fo:margin-top="' + cmFromIn(page.margins.top) + '" fo:margin-right="' + cmFromIn(page.margins.right) + '" ' +
      'fo:margin-bottom="' + cmFromIn(page.margins.bottom) + '" fo:margin-left="' + cmFromIn(page.margins.left) + '" ' +
      'style:header-height="' + cmFromIn(page.margins.header) + '" style:footer-height="' + cmFromIn(page.margins.footer) + '" ' +
      'fo:margin-header="' + cmFromIn(0.2) + '" fo:margin-footer="' + cmFromIn(0.2) + '" ' +
      'style:writing-mode="lr-tb"/></style:page-layout>' +
      '</office:automatic-styles>' +
      '</office:document-styles>';

    var meta = '<?xml version="1.0" encoding="UTF-8"?>' +
      '<office:document-meta ' + nsDecl() + '><office:meta>' +
      (opts.meta && opts.meta.title ? '<dc:title>' + U.escapeXml(opts.meta.title) + '</dc:title>' : '') +
      (opts.meta && opts.meta.subject ? '<dc:subject>' + U.escapeXml(opts.meta.subject) + '</dc:subject>' : '') +
      (opts.meta && opts.meta.author ? '<meta:initial-creator>' + U.escapeXml(opts.meta.author) + '</meta:initial-creator>' +
        '<meta:creation-date>' + new Date().toISOString().replace(/\.\d+Z$/, '') + '</meta:creation-date>' : '') +
      (opts.meta && opts.meta.keywords ? '<meta:keyword>' + U.escapeXml(opts.meta.keywords) + '</meta:keyword>' : '') +
      (opts.meta && opts.meta.comments ? '<dc:description>' + U.escapeXml(opts.meta.comments) + '</dc:description>' : '') +
      '<meta:generator>OpenWord</meta:generator>' +
      '</office:meta></office:document-meta>';

    var manifestEntries = [
      '<manifest:file-entry manifest:full-path="/" manifest:media-type="' + MIME + '"/>',
      '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>',
      '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>',
      '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>'
    ];
    w.images.forEach(function (im) {
      manifestEntries.push('<manifest:file-entry manifest:full-path="' + im.path + '" manifest:media-type="' + im.mime + '"/>');
    });
    var manifest = '<?xml version="1.0" encoding="UTF-8"?>' +
      '<manifest:manifest xmlns:manifest="' + NS.manifest + '">' + manifestEntries.join('') + '</manifest:manifest>';

    var zip = new JSZip();
    zip.file('mimetype', MIME, { compression: 'STORE' });
    zip.file('content.xml', content);
    zip.file('styles.xml', styles);
    zip.file('meta.xml', meta);
    zip.folder('META-INF').file('manifest.xml', manifest);
    w.images.forEach(function (im) { zip.folder('Pictures').file(im.path.replace('Pictures/', ''), im.bytes); });
    return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' }).then(function (blob) {
      return { blob: blob, warnings: w.warnings };
    });
  }

  function cmFromIn(v) { return (v * 2.54).toFixed(3) + 'cm'; }

  function nsDecl() {
    return 'xmlns:office="' + NS.office + '" xmlns:style="' + NS.style + '" xmlns:text="' + NS.text +
      '" xmlns:table="' + NS.table + '" xmlns:draw="' + NS.draw + '" xmlns:fo="' + NS.fo +
      '" xmlns:svg="' + NS.svg + '" xmlns:xlink="' + NS.xlink + '" xmlns:dc="' + NS.dc +
      '" xmlns:meta="' + NS.meta + '" xmlns:number="' + NS.number + '" xmlns:of="' + NS.office.replace(':office:1.0', ':of:1.0') + '"';
  }

  function hfToOdt(el, w) {
    if (!el || !U.stripZwsp(el.textContent || '').trim()) return '';
    return blocksToOdt(el, w);
  }

  /** Convert a container element's children into ODT block XML. */
  function blocksToOdt(container, w, inList) {
    var out = '';
    var pendingInline = '';
    function flush() {
      if (pendingInline) {
        out += '<text:p>' + pendingInline + '</text:p>';
        pendingInline = '';
      }
    }
    Array.prototype.forEach.call(container.childNodes, function (node) {
      if (node.nodeType === 3) { pendingInline += U.escapeXml(node.nodeValue); return; }
      if (node.nodeType !== 1) return;
      if (node.classList && node.classList.contains('ww-nosrc')) return;
      var tag = node.tagName.toLowerCase();
      if (tag === 'ul' || tag === 'ol') { flush(); out += listToOdt(node, w); return; }
      if (tag === 'table') { flush(); out += tableToOdt(node, w); return; }
      if (tag === 'hr') { flush(); out += '<text:p><text:span text:style-name="' + hrStyle(w) + '">————————————————————————————————</text:span></text:p>'; return; }
      if (tag === 'br') { pendingInline += '<text:line-break/>'; return; }
      if (node.classList && node.classList.contains('ww-pagebreak')) {
        flush();
        out += '<text:p text:style-name="' + breakStyle(w) + '"></text:p>';
        return;
      }
      if (hasBlockChildren(node)) {
        flush();
        out += blocksToOdt(node, w, inList);
        return;
      }
      if (tag === 'img') { pendingInline += imageToOdt(node, w); return; }
      flush();
      out += paraToOdt(node, w);
    });
    flush();
    if (!out) out = '<text:p/>';
    return out;
  }

  function hasBlockChildren(el) {
    return Array.prototype.some.call(el.children, function (c) {
      var t = c.tagName.toLowerCase();
      return ['p', 'div', 'ul', 'ol', 'table', 'blockquote', 'pre', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'hr', 'figure'].indexOf(t) >= 0;
    });
  }

  function paraToOdt(el, w) {
    var props = {};
    var st = el.style;
    if (st.marginLeft) props.ml = cmFromPt(U.cssLengthToPt(st.marginLeft, el) || 0);
    if (st.marginRight) props.mr = cmFromPt(U.cssLengthToPt(st.marginRight, el) || 0);
    props.mt = st.marginTop ? cmFromPt(U.cssLengthToPt(st.marginTop, el) || 0) : '0cm';
    props.mb = st.marginBottom ? cmFromPt(U.cssLengthToPt(st.marginBottom, el) || 0) : '0.247cm';
    if (st.textIndent) props.ti = cmFromPt(U.cssLengthToPt(st.textIndent, el) || 0);
    if (st.textAlign) {
      var ta = st.textAlign.toLowerCase();
      props.align = ta === 'justify' ? 'justify' : (ta === 'center' ? 'center' : ta === 'right' ? 'right' : null);
    }
    if (st.lineHeight && st.lineHeight !== 'normal') {
      if (/^[\d.]+$/.test(st.lineHeight.trim())) props.line = Math.round(parseFloat(st.lineHeight) * 100);
      else {
        var lhpt = U.cssLengthToPt(st.lineHeight, el);
        if (lhpt) props.line = 'proportional 100%';
      }
    }
    if (st.backgroundColor && !/rgba\(0, 0, 0, 0\)|transparent/i.test(st.backgroundColor)) {
      props.bg = (U.rgbStringToHex(st.backgroundColor) || '').toLowerCase();
    }
    var tag = el.tagName.toLowerCase();
    var outline = { h1: 1, h2: 2, h3: 3, h4: 4, h5: 5, h6: 6 }[tag] || 0;
    if (el.classList && el.classList.contains('ww-title')) {
      return '<text:p text:style-name="Title">' + inlineToOdt(el, w) + '</text:p>';
    }
    if (el.classList && el.classList.contains('ww-subtitle')) {
      return '<text:p text:style-name="Subtitle">' + inlineToOdt(el, w) + '</text:p>';
    }
    if (el.tagName === 'BLOCKQUOTE' || (el.classList && el.classList.contains('ww-quote'))) {
      return '<text:p text:style-name="Quotations">' + inlineToOdt(el, w) + '</text:p>';
    }
    if (el.tagName === 'PRE') {
      return '<text:p text:style-name="' + preStyle(w) + '">' + inlineToOdt(el, w, { font: 'Courier New' }) + '</text:p>';
    }
    var hasExtras = props.ml || props.mr || props.mt || props.mb || props.ti || props.align || props.line || props.bg;
    if (outline) {
      var hstyle = w.styleFor('p', Object.assign({ h: outline, extra: hasExtras ? props : null }, hasExtras ? props : {}), function (p, name) {
        var sz = { 1: '20pt', 2: '16pt', 3: '13.5pt', 4: '12pt', 5: '11pt', 6: '11pt' }[p.h];
        var extra = p.extra || {};
        return '<style:style style:name="' + name + '" style:family="paragraph" style:parent-style-name="Heading_20_' + p.h + '" style:class="text">' +
          (extra.ml ? '<style:paragraph-properties fo:margin-left="' + extra.ml + '"' : '<style:paragraph-properties') +
          (extra.mr ? ' fo:margin-right="' + extra.mr + '"' : '') +
          (extra.mt !== undefined ? ' fo:margin-top="' + extra.mt + '"' : '') +
          (extra.mb !== undefined ? ' fo:margin-bottom="' + extra.mb + '"' : '') +
          (extra.align ? ' fo:text-align="' + extra.align + '"' : '') +
          (extra.line ? ' style:line-height="' + extra.line + '"' : '') +
          '/><style:text-properties fo:font-size="' + sz + '"/></style:style>';
      });
      return '<text:h text:style-name="' + hstyle + '" text:outline-level="' + outline + '">' + inlineToOdt(el, w) + '</text:h>';
    }
    var styleName = 'Text_20_body';
    if (hasExtras) styleName = w.styleFor('p', props, paraStyleXml);
    return '<text:p text:style-name="' + styleName + '">' + inlineToOdt(el, w) + '</text:p>';
  }

  function hrStyle(w) {
    return w.styleFor('t', { hr: 1 }, function (p, name) {
      return '<style:style style:name="' + name + '" style:family="text"><style:text-properties fo:color="#8f8f8f"/></style:style>';
    });
  }
  function preStyle(w) {
    return w.styleFor('p', { pre: 1 }, function (p, name) {
      return '<style:style style:name="' + name + '" style:family="paragraph"><style:paragraph-properties fo:background-color="#f5f5f5"/>' +
        '<style:text-properties fo:font-family="Courier New" fo:font-size="10pt"/></style:style>';
    });
  }
  function breakStyle(w) {
    return w.styleFor('p', { pb: 1 }, function (p, name) {
      return '<style:style style:name="' + name + '" style:family="paragraph" style:parent-style-name="Text_20_body">' +
        '<style:paragraph-properties fo:break-before="page"/></style:style>';
    });
  }

  function listToOdt(listEl, w) {
    var ordered = listEl.tagName.toLowerCase() === 'ol';
    var fmt = ordered ? (listEl.getAttribute('data-ww-list-type') ? cssToNumFmt(listEl.getAttribute('data-ww-list-type')) : 'decimal') : null;
    w.listStyles = w.listStyles || {};
    var key = 'list' + (ordered ? 'o' : 'u') + (fmt || '') + (listEl.style.listStyleType || '');
    if (!w.listStyles[key]) {
      var name = 'L' + (Object.keys(w.listStyles).length + 1);
      w.listStyles[key] = listStyleXml({ ordered: ordered, fmt: fmt }, name);
      w.listStyles[key + '@name'] = name;
    }
    var lname = w.listStyles[key + '@name'];
    var out = '<text:list text:style-name="' + lname + '">';
    Array.prototype.forEach.call(listEl.children, function (li) {
      if (li.tagName !== 'LI') return;
      out += '<text:list-item>' + blocksToOdt(li, w, true) + '</text:list-item>';
    });
    out += '</text:list>';
    return out;
  }

  function cssToNumFmt(css) {
    return ({
      'lower-alpha': 'lowerLetter', 'upper-alpha': 'upperLetter', 'lower-roman': 'lowerRoman',
      'upper-roman': 'upperRoman', 'decimal-leading-zero': 'decimalZero', 'none': 'none'
    })[css] || 'decimal';
  }

  function inlineToOdt(el, w, baseProps) {
    var out = '';
    function walk(node, props) {
      if (node.nodeType === 3) { out += U.escapeXml(node.nodeValue); return; }
      if (node.nodeType !== 1) return;
      var tag = node.tagName.toLowerCase();
      var p = Object.assign({}, props);
      switch (tag) {
        case 'b': case 'strong': p.b = 1; break;
        case 'i': case 'em': p.i = 1; break;
        case 'u': p.u = 1; break;
        case 's': case 'strike': case 'del': p.s = 1; break;
        case 'sub': p.va = 'subscript'; break;
        case 'sup': p.va = 'superscript'; break;
        case 'mark': p.bg = '#ffff00'; break;
        case 'br': out += '<text:line-break/>'; return;
        case 'a': {
          var href = node.getAttribute('href') || '';
          if (href && href.charAt(0) !== '#') {
            out += '<text:a xlink:type="simple" xlink:href="' + U.escapeXml(href) + '" xlink:show="new">';
            Array.prototype.forEach.call(node.childNodes, function (c) { walk(c, Object.assign({ link: 1 }, p)); });
            out += '</text:a>';
            return;
          }
          break;
        }
        case 'img': out += imageToOdt(node, w); return;
        case 'span': case 'font': case 'code': case 'kbd': case 'samp': {
          if (node.classList && node.classList.contains('ww-field')) {
            var kind = node.getAttribute('data-ww-field');
            var fieldTag = { page: 'text:page-number', numpages: 'text:page-count', date: 'text:date', time: 'text:time' }[kind];
            if (fieldTag) {
              out += '<' + fieldTag + ' text:select-page="current">' + U.escapeXml(node.textContent || '1') + '</' + fieldTag + '>';
              return;
            }
          }
          var st = node.style;
          if (st.fontWeight && /^(bold|[6-9]00)$/.test(st.fontWeight)) p.b = 1;
          if (st.fontStyle === 'italic') p.i = 1;
          if (st.textDecorationLine && st.textDecorationLine.indexOf('underline') >= 0) p.u = 1;
          if (st.textDecoration && st.textDecoration.indexOf('underline') >= 0) p.u = 1;
          if (st.textDecorationLine && st.textDecorationLine.indexOf('line-through') >= 0) p.s = 1;
          if (st.color) p.color = (U.rgbStringToHex(st.color) || '').toLowerCase();
          if (st.backgroundColor && !/rgba\(0, 0, 0, 0\)|transparent/i.test(st.backgroundColor)) p.bg = (U.rgbStringToHex(st.backgroundColor) || '').toLowerCase();
          if (st.fontSize) { var pt = U.cssLengthToPt(st.fontSize, node); if (pt) p.sz = Math.round(pt * 100) / 100; }
          if (st.fontFamily) p.font = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
          if (st.verticalAlign === 'super' || st.verticalAlign === 'superscript') p.va = 'superscript';
          if (st.verticalAlign === 'sub' || st.verticalAlign === 'subscript') p.va = 'subscript';
          if (tag === 'code' || tag === 'kbd' || tag === 'samp') p.font = p.font || 'Courier New';
          break;
        }
        default: break;
      }
      var styleName = '';
      var sig = [p.b, p.i, p.u, p.s, p.color, p.bg, p.sz, p.font, p.va, p.link].join('|');
      var hasAny = /[^|]/.test(sig.replace(/\|/g, ''));
      if (hasAny) styleName = w.styleFor('t', p, textStyleXml);
      if (styleName) out += '<text:span text:style-name="' + styleName + '">';
      Array.prototype.forEach.call(node.childNodes, function (c) { walk(c, Object.assign({}, p)); });
      if (styleName) out += '</text:span>';
    }
    Array.prototype.forEach.call(el.childNodes, function (c) { walk(c, baseProps || {}); });
    return out || '<text:span></text:span>';
  }

  function imageToOdt(img, w) {
    var src = img.getAttribute('src') || '';
    var pxW = img.width || parseFloat(img.getAttribute('width')) || 0;
    var pxH = img.height || parseFloat(img.getAttribute('height')) || 0;
    if (!pxW || !pxH) { w.warnings.push('An image was skipped (no size).'); return ''; }
    var path, bytes, mime;
    if (/^data:/i.test(src)) {
      try {
        var comma = src.indexOf(',');
        mime = (/^data:([^;,]+)/.exec(src.slice(0, comma)) || [0, 'image/png'])[1];
        bytes = U.dataUrlToBytes(src);
        path = 'Pictures/img' + (w.images.length + 1) + '.' + U.extFromMime(mime);
      } catch (e) { w.warnings.push('Could not embed an image.'); return ''; }
    } else {
      w.warnings.push('Remote images are not embedded in ODT exports.');
      return '<text:span>' + U.escapeXml(img.getAttribute('alt') || src) + '</text:span>';
    }
    var dup = w.imageMap[src];
    if (!dup) {
      w.imageMap[src] = path;
      w.images.push({ path: path, bytes: bytes, mime: mime });
      dup = path;
    }
    return '<draw:frame draw:style-name="fr' + (w.images.length) + '" draw:name="' + U.escapeXml(path.split('/').pop()) +
      '" text:anchor-type="as-char" svg:width="' + cmFromPx(pxW) + '" svg:height="' + cmFromPx(pxH) + '">' +
      '<draw:image xlink:href="' + dup + '" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>';
  }

  return { parseOdt: parseOdt, exportOdt: exportOdt, MIME: MIME };
})(WW.util);
