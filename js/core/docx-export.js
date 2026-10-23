/* ==========================================================================
   OpenWord — DOCX (OOXML wordprocessingml) writer
   Hand-rolled: no external dependency besides JSZip for the zip container.
   Element order follows the ECMA-376 sequences because Word refuses to open
   packages whose children are out of order.
   ========================================================================== */
WW.docxExport = (function (U) {
  'use strict';

  var NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var NS_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
  var CT_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
  var PAGE_SIZES = {
    Letter: [12240, 15840], A4: [11906, 16838], Legal: [12240, 20160],
    A3: [16838, 23811], A5: [5953, 8419], Tabloid: [15840, 12240], B5: [10319, 14598]
  };

  /* =========================================================== context === */
  function makeContext() {
    return {
      media: [],            // {id:'media/image1.png', bytes, ext, mime}
      rels: [],             // {id:'rId3', type, target, targetMode}
      relSeq: 1,
      drawings: 0,
      numbering: { abstract: [], nums: [] },
      fonts: { Calibri: 1 },
      warnings: []
    };
  }

  function addRel(ctx, type, target, mode) {
    var id = 'rId' + (ctx.relSeq++);
    ctx.rels.push({ id: id, type: type, target: target, mode: mode || null });
    return id;
  }

  function relIdForImage(ctx, src, altText) {
    var i, existing;
    if (/^data:/i.test(src)) {
      var comma = src.indexOf(',');
      var head = src.slice(0, comma), b64 = src.slice(comma + 1);
      var mime = (/^data:([^;,]+)/.exec(head) || [0, 'image/png'])[1].toLowerCase();
      var ext = U.extFromMime(mime);
      for (i = 0; i < ctx.media.length; i++) {
        if (ctx.media[i].src === src) return ctx.media[i].relId;
      }
      var bytes;
      try { bytes = U.dataUrlToBytes(src); }
      catch (e) { ctx.warnings.push('Could not embed image (' + (altText || 'untitled') + ')'); return null; }
      var name = 'media/image' + (ctx.media.length + 1) + '.' + (ext === 'jpeg' ? 'jpeg' : ext);
      var relId = addRel(ctx, NS_DOC + '/image', name);
      ctx.media.push({ src: src, bytes: bytes, relId: relId, name: name, mime: mime, ext: ext });
      return relId;
    }
    // Remote image: try to fetch and embed; otherwise fall back to a link.
    return { external: src };
  }

  /* ============================================================ runs ===== */
  function rprXml(rp) {
    if (!rp) return '';
    var out = '';
    if (rp.rStyle) out += '<w:rStyle w:val="' + rp.rStyle + '"/>';
    if (rp.font) {
      var f = U.escapeXml(rp.font);
      out += '<w:rFonts w:ascii="' + f + '" w:hAnsi="' + f + '" w:cs="' + f + '"/>';
    }
    if (rp.bold) out += '<w:b/><w:bCs/>';
    if (rp.italic) out += '<w:i/><w:iCs/>';
    if (rp.strike) out += '<w:strike/>';
    if (rp.color) out += '<w:color w:val="' + rp.color.slice(1) + '"/>';
    if (rp.size) { var sz = String(Math.round(rp.size * 2)); out += '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/>'; }
    if (rp.highlight) out += '<w:highlight w:val="' + rp.highlight + '"/>';
    if (rp.underline) out += '<w:u w:val="' + (rp.underline === true ? 'single' : rp.underline) + '"/>';
    if (rp.vertAlign) out += '<w:vertAlign w:val="' + rp.vertAlign + '"/>';
    return out ? '<w:rPr>' + out + '</w:rPr>' : '';
  }

  function textRun(rp, text) {
    var xml = '', chunk = '';
    for (var i = 0; i < text.length; i++) {
      var ch = text[i];
      if (ch === '\t') {
        xml += run(rp, chunk); chunk = '';
        xml += '<w:r>' + rprXml(rp) + '<w:tab/></w:r>';
      } else if (ch === '\n') {
        xml += run(rp, chunk); chunk = '';
        xml += '<w:r>' + rprXml(rp) + '<w:br/></w:r>';
      } else chunk += ch;
    }
    xml += run(rp, chunk);
    return xml;
  }

  function run(rp, text) {
    if (!text) return '';
    return '<w:r>' + rprXml(rp) + '<w:t xml:space="preserve">' + U.escapeXml(text) + '</w:t></w:r>';
  }

  function cloneProps(rp) {
    var o = {}; for (var k in rp) if (Object.prototype.hasOwnProperty.call(rp, k)) o[k] = rp[k];
    return o;
  }

  function propsFromStyle(el, rp) {
    var st = el.style, fw, fs;
    if (st.fontFamily) {
      var fam = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (fam && fam !== 'inherit') { rp.font = fam; ctxFont(fam); }
    }
    if (st.fontSize) {
      var pt = U.cssLengthToPt(st.fontSize, el);
      if (pt) rp.size = pt;
    }
    fw = st.fontWeight;
    if (fw) {
      if (fw === 'bold' || fw === 'bolder' || parseInt(fw, 10) >= 600) rp.bold = true;
      else if (fw === 'normal' || parseInt(fw, 10) < 600) rp.bold = false;
    }
    if (st.fontStyle === 'italic') rp.italic = true;
    var deco = (st.textDecorationLine || st.textDecoration || '').toLowerCase();
    if (deco.indexOf('underline') >= 0) rp.underline = 'single';
    if (deco.indexOf('line-through') >= 0) rp.strike = true;
    if (deco.indexOf('none') >= 0) { rp.underline = false; rp.strike = false; }
    if (st.color) { var c = U.rgbStringToHex(st.color); if (c) rp.color = c; }
    if (st.backgroundColor) { var bg = U.rgbStringToHex(st.backgroundColor); if (bg && !isTransparent(el.style.backgroundColor)) rp.highlight = U.hexToHighlightName(bg) || undefined; if (bg && rp.highlight === undefined) rp.shading = bg; }
    if (st.verticalAlign === 'sub' || st.verticalAlign === 'subscript') rp.vertAlign = 'subscript';
    if (st.verticalAlign === 'super' || st.verticalAlign === 'superscript') rp.vertAlign = 'superscript';
    return rp;
  }

  function isTransparent(value) {
    if (!value) return true;
    var m = /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?\s*\)/i.exec(value);
    if (m && m[4] !== undefined && parseFloat(m[4]) === 0) return true;
    return /^(transparent)$/i.test(value.trim());
  }

  function ctxFont(name) { /* font table is generated from the used set */ }

  var IMG_SIZES = {};   // filled in by exportDocument() with pre-measured dimensions

  function inlineXml(node, rp, ctx) {
    var out = '';
    for (var i = 0; i < node.childNodes.length; i++) {
      out += inlineNode(node.childNodes[i], rp, ctx);
    }
    return out;
  }

  function inlineNode(node, rp, ctx) {
    if (node.nodeType === 3) return textRun(rp, node.nodeValue);
    if (node.nodeType !== 1) return '';
    if (node.classList && node.classList.contains('ww-nosrc')) return '';
    var tag = node.tagName.toLowerCase();

    switch (tag) {
      case 'br': return '<w:r>' + rprXml(rp) + '<w:br/></w:r>';
      case 'b': case 'strong': return inlineXml(node, assign(rp, { bold: true }), ctx);
      case 'i': case 'em': case 'cite': case 'var': return inlineXml(node, assign(rp, { italic: true }), ctx);
      case 'u': case 'ins': return inlineXml(node, assign(rp, { underline: 'single' }), ctx);
      case 's': case 'strike': case 'del': return inlineXml(node, assign(rp, { strike: true }), ctx);
      case 'sub': return inlineXml(node, assign(rp, { vertAlign: 'subscript' }), ctx);
      case 'sup': return inlineXml(node, assign(rp, { vertAlign: 'superscript' }), ctx);
      case 'mark': return inlineXml(node, assign(rp, { highlight: 'yellow' }), ctx);
      case 'code': case 'kbd': case 'samp': case 'tt': {
        var mono = cloneProps(rp); mono.font = mono.font || 'Consolas'; ctxFont(mono.font);
        return inlineXml(node, mono, ctx);
      }
      case 'small': { var sm = cloneProps(rp); if (sm.size) sm.size = Math.max(6, sm.size * 0.8); return inlineXml(node, sm, ctx); }
      case 'a': {
        var href = node.getAttribute('href') || '';
        var inner = inlineXml(node, assign(cloneProps(rp), { rStyle: 'Hyperlink' }), ctx);
        if (!href || href.charAt(0) === '#') return inner;
        var rid = addRel(ctx, NS_DOC + '/hyperlink', href.replace(/&amp;/g, '&'), 'External');
        return '<w:hyperlink r:id="' + rid + '">' + inner + '</w:hyperlink>';
      }
      case 'img': return drawingXml(node, ctx);
      case 'span': case 'font': case 'label': case 'abbr': case 'q': case 'output': case 'time': case 'data': case 'bdi': case 'bdo': case 'ruby': case 'nobr': {
        var nrp = propsFromStyle(node, cloneProps(rp));
        if (node.getAttribute && node.getAttribute('class') === 'ww-field') return fieldXml(node, nrp, ctx);
        var content = inlineXml(node, nrp, ctx);
        // Nested block elements inside a span: close and continue the flow.
        return content;
      }
      case 'script': case 'style': case 'noscript': case 'iframe': case 'object': case 'embed': return '';
      default:
        if (node.classList && node.classList.contains('ww-field')) return fieldXml(node, cloneProps(rp), ctx);
        // Any other element: treat as inline container.
        return inlineXml(node, propsFromStyle(node, cloneProps(rp)), ctx);
    }
  }

  function assign(rp, patch) {
    var o = cloneProps(rp);
    for (var k in patch) o[k] = patch[k];
    return o;
  }

  function fieldXml(node, rp, ctx) {
    var kind = (node.getAttribute('data-ww-field') || '').toLowerCase();
    var instr = null;
    if (kind === 'page') instr = ' PAGE   \\* MERGEFORMAT ';
    else if (kind === 'numpages') instr = ' NUMPAGES   \\* MERGEFORMAT ';
    else if (kind === 'pages') instr = ' NUMPAGES   \\* MERGEFORMAT ';
    else if (kind === 'date') instr = ' DATE   \\* MERGEFORMAT ';
    else if (kind === 'time') instr = ' TIME   \\* MERGEFORMAT ';
    else if (kind === 'filename') instr = ' FILENAME   \\* MERGEFORMAT ';
    else if (kind === 'author') instr = ' AUTHOR   \\* MERGEFORMAT ';
    if (!instr) return inlineXml(node, rp, ctx);
    return '<w:fldSimple w:instr="' + U.escapeXml(instr) + '"><w:r>' + rprXml(rp) +
      '<w:t xml:space="preserve">' + U.escapeXml(node.textContent || '1') + '</w:t></w:r></w:fldSimple>';
  }

  function drawingXml(img, ctx) {
    var src = img.getAttribute('src') || '';
    var alt = img.getAttribute('alt') || '';
    var id = ++ctx.drawings;
    var pxW = measure(img, 'width'), pxH = measure(img, 'height');
    if (!pxW || !pxH) { ctx.warnings.push('Image "' + (alt || src.slice(0, 30)) + '" has no size and was skipped.'); return ''; }
    var relId = relIdForImage(ctx, src, alt);
    if (!relId) return '';
    if (relId.external) {
      ctx.warnings.push('Remote image not embedded (browser blocked the download); inserted as a link instead.');
      var rid = addRel(ctx, NS_DOC + '/hyperlink', relId.external, 'External');
      return '<w:hyperlink r:id="' + rid + '">' + run({ rStyle: 'Hyperlink', italic: true }, alt || relId.external) + '</w:hyperlink>';
    }
    var cx = U.pxToEmu(pxW), cy = U.pxToEmu(pxH);
    var name = 'Picture ' + id + (alt ? ' (' + alt.replace(/"/g, '') + ')' : '');
    return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:docPr id="' + id + '" name="' + U.escapeXml(name) + '"/>' +
      '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="' + id + '" name="' + U.escapeXml(name) + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + relId + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  }

  function measure(img, dim) {
    var key = dim === 'width' ? 'w' : 'h';
    var cacheKey = img.getAttribute('src') || img.dataset.wwKey;
    if (IMG_SIZES[cacheKey] && IMG_SIZES[cacheKey][key]) return IMG_SIZES[cacheKey][key];
    var v = img[dim];
    if (!v || v < 2) {
      var attr = img.getAttribute(dim);
      v = attr ? parseFloat(attr) : 0;
    }
    return v || 0;
  }

  /* ======================================================== paragraphs === */
  function pPrXml(pp) {
    var out = '';
    if (pp.style) out += '<w:pStyle w:val="' + pp.style + '"/>';
    if (pp.pageBreakBefore) out += '<w:pageBreakBefore/>';
    if (pp.numPr) {
      out += '<w:numPr><w:ilvl w:val="' + pp.numPr.ilvl + '"/><w:numId w:val="' + pp.numPr.numId + '"/></w:numPr>';
    }
    if (pp.shd) out += '<w:shd w:val="clear" w:color="auto" w:fill="' + pp.shd.slice(1) + '"/>';
    if (pp.tabs && pp.tabs.length) {
      out += '<w:tabs>' + pp.tabs.map(function (t) {
        return '<w:tab w:val="' + t.val + '" w:leader="' + t.leader + '" w:pos="' + t.pos + '"/>';
      }).join('') + '</w:tabs>';
    }
    if (pp.before !== undefined || pp.after !== undefined || pp.line !== undefined) {
      out += '<w:spacing' +
        (pp.before !== undefined ? ' w:before="' + pp.before + '"' : '') +
        (pp.after !== undefined ? ' w:after="' + pp.after + '"' : '') +
        (pp.line !== undefined ? ' w:line="' + pp.line + '" w:lineRule="' + (pp.lineRule || 'auto') + '"' : '') +
        '/>';
    }
    if (pp.left !== undefined || pp.right !== undefined || pp.firstLine !== undefined || pp.hanging !== undefined) {
      out += '<w:ind' +
        (pp.left !== undefined ? ' w:left="' + pp.left + '"' : '') +
        (pp.right !== undefined ? ' w:right="' + pp.right + '"' : '') +
        (pp.firstLine !== undefined ? ' w:firstLine="' + pp.firstLine + '"' : '') +
        (pp.hanging !== undefined ? ' w:hanging="' + pp.hanging + '"' : '') +
        '/>';
    }
    if (pp.jc) out += '<w:jc w:val="' + pp.jc + '"/>';
    if (pp.outlineLvl !== undefined) out += '<w:outlineLvl w:val="' + pp.outlineLvl + '"/>';
    return out ? '<w:pPr>' + out + '</w:pPr>' : '';
  }

  var JC = { left: 'left', center: 'center', right: 'right', justify: 'both', start: 'left', end: 'right', initial: 'left', matchparent: 'left' };

  /** Read block-level properties from an element's inline styles. */
  function readBlockProps(el, defaults) {
    var pp = {}, st = el.style;
    if (st.textAlign && JC[st.textAlign.toLowerCase()]) pp.jc = JC[st.textAlign.toLowerCase()];
    var ml = cssPx(st.marginLeft), mr = cssPx(st.marginRight);
    if (ml) pp.left = U.ptToTwip(ml);
    if (mr) pp.right = U.ptToTwip(mr);
    var ti = st.textIndent ? U.cssLengthToPt(st.textIndent, el) : null;
    if (ti) { if (ti < 0) pp.hanging = Math.round(-ti * 20); else pp.firstLine = Math.round(ti * 20); }
    var mt = cssPx(st.marginTop), mb = cssPx(st.marginBottom);
    if (mt !== null) pp.before = Math.round(mt * 20 / (96 / 72));
    if (mb !== null) pp.after = Math.round(mb * 20 / (96 / 72));
    if (st.lineHeight && st.lineHeight !== 'normal') {
      var lh = parseFloat(st.lineHeight);
      var isUnitless = /^[\d.]+$/.test(String(st.lineHeight).trim());
      if (isUnitless) {
        pp.line = Math.round(lh * 240);
        pp.lineRule = 'auto';
      } else {
        var lhPt = U.cssLengthToPt(st.lineHeight, el);
        if (lhPt) { pp.line = Math.round(lhPt * 20); pp.lineRule = 'exact'; }
      }
    }
    if (st.backgroundColor && !isTransparent(st.backgroundColor)) {
      var bg = U.rgbStringToHex(st.backgroundColor);
      if (bg) pp.shd = bg;
    }
    return pp;
  }

  function cssPx(v) {
    if (!v) return null;
    var m = /^(-?[\d.]+)px$/.exec(String(v).trim());
    return m ? parseFloat(m[1]) : (/^0$/.test(String(v).trim()) ? 0 : null);
  }

  var HEADING_STYLES = { h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', h4: 'Heading4', h5: 'Heading5', h6: 'Heading6' };
  var CLASS_STYLES = {
    'ww-title': 'Title', 'ww-subtitle': 'Subtitle', 'ww-quote': 'Quote',
    'ww-intense-quote': 'IntenseQuote', 'ww-list-paragraph': 'ListParagraph', 'ww-no-spacer': null
  };

  function styleForBlock(el) {
    if (el.classList) {
      for (var i = 0; i < el.classList.length; i++) {
        var c = el.classList[i];
        if (Object.prototype.hasOwnProperty.call(CLASS_STYLES, c)) return CLASS_STYLES[c];
        if (/^ww-/.test(c)) continue;
      }
    }
    if (el.tagName === 'BLOCKQUOTE') return 'Quote';
    return HEADING_STYLES[el.tagName.toLowerCase()] || null;
  }

  /* ============================================================= body ==== */
  function hasBlockChild(el) {
    for (var i = 0; i < el.children.length; i++) {
      var c = el.children[i];
      var t = c.tagName.toLowerCase();
      if (['p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'table', 'blockquote', 'pre', 'figure', 'section', 'hr', 'header', 'footer'].indexOf(t) >= 0) return true;
    }
    return false;
  }

  function isInlineLevel(el) {
    var t = el.tagName.toLowerCase();
    return ['img', 'br', 'hr'].indexOf(t) >= 0 || (!hasBlockChild(el) &&
      ['table', 'ul', 'ol', 'div', 'p'].indexOf(t) < 0);
  }

  /** Turn a container element's children into a sequence of body-level items. */
  function collectBlocks(container, ctx) {
    var out = [];
    var pending = null;
    function flushPending() { if (pending) { out.push(pending); pending = null; } }
    Array.prototype.forEach.call(container.childNodes, function (node) {
      if (node.nodeType === 3) {
        if (!U.stripZwsp(node.nodeValue).trim() && !/\u00a0/.test(node.nodeValue)) {
          if (node.nodeValue && node.nodeValue.length && pending) pending.el.appendChild(node.cloneNode(true));
          return;
        }
        if (!pending) { pending = makeVirtualPara(container); }
        pending.el.appendChild(node.cloneNode(true));
        return;
      }
      if (node.nodeType !== 1) return;
      var el = node, tag = el.tagName.toLowerCase();
      if (el.classList && el.classList.contains('ww-nosrc')) return;
      if (tag === 'ul' || tag === 'ol') { flushPending(); out.push({ type: 'list', el: el }); return; }
      if (tag === 'table') { flushPending(); out.push({ type: 'table', el: el }); return; }
      if (tag === 'hr') { flushPending(); out.push({ type: 'hr', el: el }); return; }
      if (el.classList && el.classList.contains('ww-pagebreak')) { flushPending(); out.push({ type: 'pagebreak' }); return; }
      if (hasBlockChild(el)) {
        // A wrapper div — recurse but keep an eventual page break marker.
        var nested = collectBlocks(el, ctx);
        if (el.classList && el.classList.contains('ww-pagebreak')) nested.push({ type: 'pagebreak' });
        nested.forEach(function (item) { out.push(item); });
        return;
      }
      if (tag === 'img' || tag === 'br') {
        if (!pending) pending = makeVirtualPara(container);
        pending.el.appendChild(el.cloneNode(true));
        return;
      }
      flushPending();
      out.push({ type: 'para', el: el });
    });
    flushPending();
    if (!out.length) out.push({ type: 'para', el: makeVirtualPara(container) });
    return out;
  }

  function makeVirtualPara(container) {
    var p = container.ownerDocument.createElement('p');
    p.dataset.virtual = '1';
    return p;
  }

  function paraXml(el, ctx, opts) {
    opts = opts || {};
    var pp = {};
    var style = styleForBlock(el);
    if (opts.style) style = opts.style;
    if (style) pp.style = style;
    var inlineProps = readBlockProps(el, {});
    Object.keys(inlineProps).forEach(function (k) { pp[k] = inlineProps[k]; });
    if (opts.jc !== undefined && inlineProps.jc === undefined) pp.jc = opts.jc;
    if (opts.numPr) pp.numPr = opts.numPr;
    if (opts.keepNext) pp.keepNext = true;
    var inner = inlineXml(el, {}, ctx);
    if (opts.trailingPageBreak) inner += '<w:r><w:br w:type="page"/></w:r>';
    if (!inner) inner = '';
    return '<w:p>' + pPrXml(pp) + inner + '</w:p>';
  }

  function listXml(listEl, ctx, numId, ilvl, inherited) {
    var out = '';
    var id = inherited || allocateNumId(ctx, listEl);
    Array.prototype.forEach.call(listEl.children, function (li) {
      if (li.tagName !== 'LI') return;
      var kids = Array.prototype.slice.call(li.childNodes);
      var inlineParts = [], blockParts = [], subLists = [];
      kids.forEach(function (n) {
        var t = n.nodeType === 1 ? n.tagName.toLowerCase() : null;
        if (t === 'ul' || t === 'ol') subLists.push(n);
        else if (t && ['p', 'div', 'table', 'blockquote', 'pre', 'h1', 'h2', 'h3', 'h4'].indexOf(t) >= 0 && hasBlockChild(n)) blockParts.push(n);
        else if (t && ['table'].indexOf(t) >= 0) blockParts.push(n);
        else inlineParts.push(n);
      });
      var holder = li.ownerDocument.createElement('p');
      inlineParts.forEach(function (n) { holder.appendChild(n.cloneNode(true)); });
      var jcVal = null;
      out += paraXml(holder, ctx, {
        numPr: { ilvl: ilvl, numId: id },
        jc: jcVal,
        style: 'ListParagraph'
      });
      blockParts.forEach(function (bp) {
        var t = bp.tagName.toLowerCase();
        if (t === 'table') out += tableXml(bp, ctx);
        else if (hasBlockChild(bp)) collectBlocks(bp, ctx).forEach(function (item) { out += renderItem(item, ctx); });
        else out += paraXml(bp, ctx, {});
      });
      subLists.forEach(function (sub) { out += listXml(sub, ctx, id, ilvl + 1, id); });
    });
    return out;
  }

  function fmtForList(el) {
    var tag = el.tagName.toLowerCase();
    var attr = el.getAttribute('data-ww-list-type');
    if (tag !== 'ol') return 'bullet';
    if (attr) return attr;
    var st = el.style && el.style.listStyleType;
    var map = {
      loweralpha: 'lowerLetter', lowerlatin: 'lowerLetter', upperalpha: 'upperLetter',
      upperlatin: 'upperLetter', lowerroman: 'lowerRoman', upperroman: 'upperRoman',
      decimalleadingzero: 'decimalZero', none: 'none', decimal: 'decimal'
    };
    if (st) {
      var key = st.replace(/-/g, '').toLowerCase();
      if (map[key]) return map[key];
    }
    return 'decimal';
  }

  var ABSTRACT_DEFS = {
    bullet: { fmt: 'bullet', marks: ['\u2022', '\u25E6', '\u25AA', '\u2022', '\u25E6', '\u25AA', '\u2022', '\u25E6', '\u25AA'], font: 'Arial' },
    decimal: { fmt: 'decimal', suffix: '.' },
    lowerLetter: { fmt: 'lowerLetter', suffix: '.' },
    upperLetter: { fmt: 'upperLetter', suffix: '.' },
    lowerRoman: { fmt: 'lowerRoman', suffix: '.' },
    upperRoman: { fmt: 'upperRoman', suffix: '.' },
    decimalZero: { fmt: 'decimalZero', suffix: '.' },
    none: { fmt: 'none', suffix: '' }
  };

  function abstractIdFor(ctx, fmt) {
    var def = ABSTRACT_DEFS[fmt] || ABSTRACT_DEFS.decimal;
    for (var i = 0; i < ctx.numbering.abstract.length; i++) {
      if (ctx.numbering.abstract[i].fmt === fmt) return i;
    }
    ctx.numbering.abstract.push({ fmt: fmt, def: def, nsid: randomNsid() });
    return ctx.numbering.abstract.length - 1;
  }

  function randomNsid() {
    var s = '';
    for (var i = 0; i < 8; i++) s += '0123456789ABCDEF'[Math.floor(Math.random() * 16)];
    return s;
  }

  function allocateNumId(ctx, listEl) {
    var fmt = fmtForList(listEl);
    var abs = abstractIdFor(ctx, fmt);
    var numId = ctx.numbering.nums.length + 1;
    ctx.numbering.nums.push({ numId: numId, abstractId: abs });
    return numId;
  }

  function numberingXml(ctx) {
    if (!ctx.numbering.nums.length) {
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="' + NS_W + '"/>';
    }
    var abs = '';
    ctx.numbering.abstract.forEach(function (a, idx) {
      var def = a.def, lv = '';
      for (var i = 0; i < 9; i++) {
        var lvlText = def.marks ? def.marks[i] : ('%' + (i + 1) + (def.suffix || '.'));
        var ind = ' w:left="' + (720 * (i + 1)) + '" w:hanging="360"';
        lv += '<w:lvl w:ilvl="' + i + '">' +
          '<w:start w:val="1"/>' +
          '<w:numFmt w:val="' + def.fmt + '"/>' +
          '<w:pPr><w:ind' + ind + '/><w:contextualSpacing/></w:pPr>' +
          (def.font ? '<w:rPr><w:rFonts w:ascii="' + def.font + '" w:hAnsi="' + def.font + '" w:hint="default"/></w:rPr>' : '') +
          '<w:lvlText w:val="' + U.escapeXml(lvlText) + '"/>' +
          '<w:lvlJc w:val="left"/>' +
          '</w:lvl>';
      }
      abs += '<w:abstractNum w:abstractNumId="' + idx + '">' +
        '<w:nsid w:val="' + a.nsid + '"/>' +
        '<w:multiLevelType w:val="hybridMultilevel"/>' +
        lv + '</w:abstractNum>';
    });
    var nums = ctx.numbering.nums.map(function (n) {
      var ov = '';
      for (var i = 0; i < 9; i++) ov += '<w:lvlOverride w:ilvl="' + i + '"><w:startOverride w:val="1"/></w:lvlOverride>';
      return '<w:num w:numId="' + n.numId + '"><w:abstractNumId w:val="' + n.abstractId + '"/>' + ov + '</w:num>';
    }).join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering xmlns:w="' + NS_W + '">' + abs + nums + '</w:numbering>';
  }

  function renderItem(item, ctx) {
    if (item.type === 'list') return listXml(item.el, ctx, null, 0, null);
    if (item.type === 'table') return tableXml(item.el, ctx);
    if (item.type === 'hr') return hrXml();
    if (item.type === 'pagebreak') return '<w:p>' + pPrXml({}) + '<w:r><w:br w:type="page"/></w:r></w:p>';
    return paraXml(item.el, ctx, {});
  }

  function hrXml() {
    // Horizontal rule: a paragraph with a bottom border.
    return '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="8F8F8F"/></w:pBdr>' +
      '<w:spacing w:before="120" w:after="120"/></w:pPr></w:p>';
  }

  /* =========================================================== tables === */
  function tableXml(table, ctx) {
    var rowsEls = [];
    Array.prototype.forEach.call(table.rows || [], function (r) { rowsEls.push(r); });
    if (!rowsEls.length) return '';
    // Build occupancy grid to expand colSpan / rowSpan.
    var cols = 0;
    rowsEls.forEach(function (r) {
      var sum = 0;
      Array.prototype.forEach.call(r.cells, function (c) { sum += c.colSpan || 1; });
      cols = Math.max(cols, sum);
    });
    if (!cols) return '';
    var contentWidth = ctx.page.contentWidthTwips;
    var widths = columnWidths(table, rowsEls, cols, contentWidth);
    var grid = '<w:tblGrid>' + widths.map(function (w) { return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
    var total = widths.reduce(function (a, b) { return a + b; }, 0);

    var occupied = {};   // "row,col" -> {fromAbove:true}
    var trXml = '';
    rowsEls.forEach(function (row, ri) {
      var col = 0, cells = '';
      var rowCells = Array.prototype.slice.call(row.cells);
      for (var ci = 0; ci < cols;) {
        var key = ri + ',' + ci;
        if (occupied[key]) {
          cells += '<w:tc><w:tcPr><w:tcW w:w="' + widths[ci] + '" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>';
          ci++;
          continue;
        }
        var cell = null;
        for (var k = 0; k < rowCells.length; k++) {
          if (!rowCells[k]._used) { cell = rowCells[k]; break; }
        }
        if (!cell) {
          cells += '<w:tc><w:tcPr><w:tcW w:w="' + widths[ci] + '" w:type="dxa"/></w:tcPr><w:p/></w:tc>';
          ci++;
          continue;
        }
        cell._used = true;
        var cs = cell.colSpan || 1, rs = cell.rowSpan || 1;
        cs = Math.min(cs, cols - ci);
        var w = 0;
        for (var q = 0; q < cs; q++) { w += widths[ci + q]; }
        for (var rr = ri; rr < ri + rs; rr++) {
          for (var cc = ci; cc < ci + cs; cc++) {
            if (rr !== ri || cc !== ci) occupied[rr + ',' + cc] = true;
          }
        }
        cells += cellXml(cell, ctx, w, cs, rs, row);
        ci += cs;
      }
      var trPr = '';
      if (ri === 0 && row.cells.length && table.getAttribute('data-header') === 'true') {
        trPr = '<w:trPr><w:tblHeader/></w:trPr>';
      }
      trXml += '<w:tr>' + trPr + cells + '</w:tr>';
    });
    rowsEls.forEach(function (r) { Array.prototype.forEach.call(r.cells, function (c) { delete c._used; }); });

    var tblPr = '<w:tblPr>' +
      '<w:tblStyle w:val="TableGrid"/>' +
      '<w:tblW w:w="' + total + '" w:type="dxa"/>' +
      tblBordersXml(table) +
      '<w:tblLayout w:type="fixed"/>' +
      '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>' +
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
      '</w:tblPr>';
    return '<w:tbl>' + tblPr + grid + trXml + '</w:tbl>';
  }

  function tblBordersXml(table) {
    var st = table.style || {};
    var visible = !(st.borderStyle === 'none' || table.getAttribute('data-borders') === 'none');
    if (!visible) {
      return '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (s) {
        return '<w:' + s + ' w:val="none" w:sz="0" w:space="0" w:color="auto"/>';
      }).join('') + '</w:tblBorders>';
    }
    var color = U.rgbStringToHex(st.borderTopColor || st.borderColor) || '#8F8F8F';
    var sz = 4;
    var bw = parseFloat((st.borderTopWidth || '').replace('px', ''));
    if (bw) sz = Math.max(2, Math.round(bw * 8));
    return '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (s) {
      return '<w:' + s + ' w:val="single" w:sz="' + sz + '" w:space="0" w:color="' + color.slice(1) + '"/>';
    }).join('') + '</w:tblBorders>';
  }

  function columnWidths(table, rowsEls, cols, contentWidth) {
    var out = new Array(cols), total = 0, anyExplicit = false;
    var colsEls = table.querySelectorAll(':scope > colgroup > col');
    if (colsEls.length) {
      for (var i = 0; i < cols; i++) {
        var c = colsEls[i];
        var wpx = c ? (parseFloat(c.getAttribute('width')) || 0) : 0;
        if (!wpx && c && c.style.width) wpx = parseFloat(c.style.width) || 0;
        out[i] = wpx ? Math.round(wpx * 15) : 0;   // 1px = 15 twips
        if (out[i]) { total += out[i]; anyExplicit = true; }
      }
    }
    if (!anyExplicit && rowsEls.length) {
      var first = rowsEls[0];
      var idx = 0;
      Array.prototype.forEach.call(first.cells, function (cell) {
        var cs = cell.colSpan || 1;
        var wpx = parseFloat(cell.style.width) || cell.getBoundingClientRect().width || 0;
        var twips = wpx ? Math.round(wpx * 15) : 0;
        for (var q = 0; q < cs && idx + q < cols; q++) {
          out[idx + q] = twips ? Math.round(twips / cs) : 0;
        }
        idx += cs;
      });
      anyExplicit = out.some(function (w) { return !!w; });
    }
    if (!anyExplicit) {
      var even = Math.floor(contentWidth / cols);
      out = new Array(cols).fill(even);
      out[cols - 1] = contentWidth - even * (cols - 1);
      return out;
    }
    // Scale the explicit widths so that the table keeps the width it had on screen.
    var tableW = parseFloat(table.style.width) || 0;
    var target = tableW ? Math.round(tableW * 15) : contentWidth;
    if (total <= 0) total = target;
    var scale = target / total;
    var acc = 0;
    for (var j = 0; j < cols; j++) {
      out[j] = out[j] ? Math.max(180, Math.round(out[j] * scale)) : 0;
      acc += out[j];
    }
    var zeroCols = cols - out.filter(function (w) { return w > 0; }).length;
    if (zeroCols > 0) {
      var share = Math.max(180, Math.round((target - acc) / Math.max(1, zeroCols)));
      for (var z = 0; z < cols; z++) if (!out[z]) out[z] = share;
      acc = out.reduce(function (a, b) { return a + b; }, 0);
    }
    out[cols - 1] += Math.max(0, target - acc);
    return out;
  }

  function cellXml(cell, ctx, widthTwips, colSpan, rowSpan, row) {
    var tcPr = '<w:tcW w:w="' + widthTwips + '" w:type="dxa"/>' +
      (colSpan > 1 ? '<w:gridSpan w:val="' + colSpan + '"/>' : '') +
      (rowSpan > 1 ? '<w:vMerge w:val="restart"/>' : '');
    var st = cell.style || {};
    var borderStr = bordersForCell(st);
    var shd = '';
    if (st.backgroundColor && !isTransparent(st.backgroundColor)) {
      var bg = U.rgbStringToHex(st.backgroundColor);
      if (bg) shd = '<w:shd w:val="clear" w:color="auto" w:fill="' + bg.slice(1) + '"/>';
    }
    var valign = { top: 'top', middle: 'center', center: 'center', bottom: 'bottom' }[st.verticalAlign || (cell.getAttribute('valign') || '')] || '';
    tcPr += borderStr + shd + (valign ? '<w:vAlign w:val="' + valign + '"/>' : '');
    tcPr = tcPr ? '<w:tcPr>' + tcPr + '</w:tcPr>' : '';

    var cellJc = JC[(st.textAlign || '').toLowerCase()] || (cell.tagName === 'TH' ? null : null);
    var items = collectBlocks(cell, ctx);
    var body = items.map(function (item) {
      if (item.type === 'para') {
        var jc = cellJc ? (readBlockProps(item.el, {}).jc || cellJc) : undefined;
        return paraXml(item.el, ctx, { jc: jc });
      }
      return renderItem(item, ctx);
    }).join('');
    if (!/<w:p[ >]/.test(body)) body = '<w:p/>' + body;
    return '<w:tc>' + tcPr + body + '</w:tc>';
  }

  function bordersForCell(st) {
    if (!st.borderTopStyle && !st.borderBottomStyle && !st.borderLeftStyle && !st.borderRightStyle) return '';
    var parts = '';
    var sides = { top: 'top', left: 'left', bottom: 'bottom', right: 'right' };
    var any = false;
    Object.keys(sides).forEach(function (k) {
      var styleKey = 'border' + k.charAt(0).toUpperCase() + k.slice(1) + 'Style';
      var wKey = 'border' + k.charAt(0).toUpperCase() + k.slice(1) + 'Width';
      var cKey = 'border' + k.charAt(0).toUpperCase() + k.slice(1) + 'Color';
      if (!st[styleKey]) return;
      any = true;
      var val = st[styleKey] === 'none' ? 'none' : (st[styleKey] === 'double' ? 'double' : 'single');
      var w = Math.max(2, Math.round((parseFloat(st[wKey]) || 1) * 8));
      var color = U.rgbStringToHex(st[cKey]) || '#8F8F8F';
      parts += '<w:' + sides[k] + ' w:val="' + val + '" w:sz="' + w + '" w:space="0" w:color="' + color.slice(1) + '"/>';
    });
    return any ? '<w:tcBorders>' + parts + '</w:tcBorders>' : '';
  }

  /* ========================================================== sections === */
  function sectionPropertiesXml(page, ctx, headerIds, footerIds) {
    var size = PAGE_SIZES[page.size] || PAGE_SIZES.Letter;
    var w = size[0], hgt = size[1];
    if (page.orientation === 'landscape') { var t = w; w = hgt; hgt = t; }
    var m = page.margins;
    var out = '';
    (headerIds || []).forEach(function (id) { out += '<w:headerReference w:type="default" r:id="' + id + '"/>'; });
    (footerIds || []).forEach(function (id) { out += '<w:footerReference w:type="default" r:id="' + id + '"/>'; });
    out += '<w:type w:val="nextPage"/>';
    out += '<w:pgSz w:w="' + w + '" w:h="' + hgt + '"' + (page.orientation === 'landscape' ? ' w:orient="landscape"' : '') + '/>';
    out += '<w:pgMar w:top="' + U.inToTwip(m.top) + '" w:right="' + U.inToTwip(m.right) +
      '" w:bottom="' + U.inToTwip(m.bottom) + '" w:left="' + U.inToTwip(m.left) +
      '" w:header="' + U.inToTwip(m.header || 0.5) + '" w:footer="' + U.inToTwip(m.footer || 0.5) + '" w:gutter="0"/>';
    out += '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/>';
    return '<w:sectPr>' + out + '</w:sectPr>';
  }

  /* =========================================================== styles ==== */
  function stylesXml(defaults) {
    var font = defaults.font || 'Calibri';
    var size = defaults.size || 11;
    var spacing = defaults.lineSpacing || 1.15;
    function st(id, name, opts) {
      opts = opts || {};
      var pPr = opts.pPr || '', rPr = opts.rPr || '';
      return '<w:style w:type="' + (opts.type || 'paragraph') + '" w:styleId="' + id + '"' + (opts.def ? ' w:default="1"' : '') + '>' +
        '<w:name w:val="' + name + '"/>' +
        (opts.basedOn ? '<w:basedOn w:val="' + opts.basedOn + '"/>' : '') +
        (opts.next ? '<w:next w:val="' + opts.next + '"/>' : '') +
        (opts.link ? '<w:link w:val="' + opts.link + '"/>' : '') +
        (opts.uiPriority !== undefined ? '<w:uiPriority w:val="' + opts.uiPriority + '"/>' : '') +
        (opts.qFormat ? '<w:qFormat/>' : '') +
        (pPr ? '<w:pPr>' + pPr + '</w:pPr>' : '') +
        (rPr ? '<w:rPr>' + rPr + '</w:rPr>' : '') +
        '</w:style>';
    }
    function rfont(name, sz, extra) {
      return '<w:rFonts w:ascii="' + name + '" w:hAnsi="' + name + '" w:cs="' + name + '"/>' +
        (sz ? '<w:sz w:val="' + (sz * 2) + '"/><w:szCs w:val="' + (sz * 2) + '"/>' : '') + (extra || '');
    }
    var out = '';
    out += st('Normal', 'Normal', {
      def: true, qFormat: true,
      pPr: '<w:spacing w:before="0" w:after="' + Math.round(8 * 20) + '" w:line="' + Math.round(spacing * 240) + '" w:lineRule="auto"/>' +
        '<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/>',
      rPr: rfont(font, size, '<w:color w:val="000000"/><w:lang w:val="en-US"/>')
    });
    out += st('Heading1', 'heading 1', {
      basedOn: 'Normal', next: 'Normal', qFormat: true, uiPriority: 9,
      pPr: '<w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="120" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="0"/>',
      rPr: rfont('Calibri Light', 20, '<w:b w:val="0"/><w:color w:val="2B579A"/>')
    });
    out += st('Heading2', 'heading 2', {
      basedOn: 'Normal', next: 'Normal', qFormat: true, uiPriority: 9,
      pPr: '<w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="100" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="1"/>',
      rPr: rfont('Calibri Light', 16, '<w:b w:val="0"/><w:color w:val="2B579A"/>')
    });
    out += st('Heading3', 'heading 3', {
      basedOn: 'Normal', next: 'Normal', qFormat: true, uiPriority: 9,
      pPr: '<w:keepNext/><w:keepLines/><w:spacing w:before="180" w:after="90" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="2"/>',
      rPr: rfont('Calibri Light', 13.5, '<w:b w:val="0"/><w:color w:val="3B5B8C"/>')
    });
    out += st('Heading4', 'heading 4', {
      basedOn: 'Normal', next: 'Normal', qFormat: true, uiPriority: 9,
      pPr: '<w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="3"/>',
      rPr: rfont(font, 12, '<w:b w:val="0"/><w:i/><w:color w:val="3B5B8C"/>')
    });
    out += st('Heading5', 'heading 5', {
      basedOn: 'Normal', next: 'Normal', qFormat: true, uiPriority: 9,
      pPr: '<w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="4"/>',
      rPr: rfont(font, 11, '<w:color w:val="3B5B8C"/>')
    });
    out += st('Heading6', 'heading 6', {
      basedOn: 'Normal', next: 'Normal', qFormat: true, uiPriority: 9,
      pPr: '<w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="5"/>',
      rPr: rfont(font, 11, '<w:i/><w:color w:val="44546A"/>')
    });
    out += st('Title', 'Title', {
      basedOn: 'Normal', next: 'Subtitle', qFormat: true, uiPriority: 10,
      pPr: '<w:spacing w:before="0" w:after="80" w:line="240" w:lineRule="auto"/><w:outlineLvl w:val="0"/>',
      rPr: rfont('Calibri Light', 26, '<w:b w:val="0"/><w:color w:val="2B579A"/>')
    });
    out += st('Subtitle', 'Subtitle', {
      basedOn: 'Normal', next: 'Normal', qFormat: true, uiPriority: 11,
      pPr: '<w:spacing w:before="0" w:after="320"/>',
      rPr: rfont(font, 13, '<w:i/><w:color w:val="5B7A9D"/>')
    });
    out += st('Quote', 'Quote', {
      basedOn: 'Normal', next: 'Normal', qFormat: true, uiPriority: 29,
      pPr: '<w:spacing w:before="80" w:after="80"/><w:ind w:left="480" w:right="480"/>',
      rPr: '<w:i/><w:color w:val="404040"/>'
    });
    out += st('IntenseQuote', 'Intense Quote', {
      basedOn: 'Normal', next: 'Normal', qFormat: true, uiPriority: 30,
      pPr: '<w:spacing w:before="120" w:after="120"/><w:ind w:left="480" w:right="480"/>' +
        '<w:pBdr><w:left w:val="single" w:sz="18" w:space="9" w:color="2B579A"/></w:pBdr>',
      rPr: '<w:i/><w:color w:val="2B579A"/>'
    });
    out += st('ListParagraph', 'List Paragraph', {
      basedOn: 'Normal', qFormat: true, uiPriority: 34,
      pPr: '<w:spacing w:before="0" w:after="80" w:line="240" w:lineRule="auto"/><w:ind w:left="720"/><w:contextualSpacing/>'
    });
    out += st('Hyperlink', 'Hyperlink', {
      type: 'character', qFormat: true,
      rPr: '<w:color w:val="0563C1"/><w:u w:val="single"/>'
    });
    out += st('TableGrid', 'Table Grid', {
      type: 'table', qFormat: true,
      pPr: ''
    });
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="' + NS_W + '">' + out + '</w:styles>';
  }

  /* ==================================================== package pieces ==== */
  function contentTypesXml(ctx, hasHeader, hasFooter) {
    var overrides = [
      ['/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'],
      ['/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'],
      ['/word/numbering.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'],
      ['/word/settings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'],
      ['/word/fontTable.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'],
      ['/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'],
      ['/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml']
    ];
    if (hasHeader) overrides.push(['/word/header1.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml']);
    if (hasFooter) overrides.push(['/word/footer1.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml']);
    ctx.media.forEach(function (m) {
      overrides.push(['/word/' + m.name, m.mime === 'image/svg+xml' ? 'image/svg+xml' : U.mimeFromExt(m.ext)]);
    });
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="' + CT + '">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
      '<Default Extension="jpg" ContentType="image/jpg"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Default Extension="gif" ContentType="image/gif"/>' +
      '<Default Extension="webp" ContentType="image/webp"/>' +
      '<Default Extension="svg" ContentType="image/svg+xml"/>' +
      '<Default Extension="bmp" ContentType="image/bmp"/>' +
      overrides.map(function (o) { return '<Override PartName="' + o[0] + '" ContentType="' + o[1] + '"/>'; }).join('') +
      '</Types>';
  }

  function rootRelsXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="' + CT_REL + '">' +
      '<Relationship Id="rId1" Type="' + NS_DOC + '/officeDocument" Target="word/document.xml"/>' +
      '<Relationship Id="rId2" Type="' + NS_DOC.replace('officeDocument/2006/relationships', 'package/2006/relationships/metadata/core-properties') + '" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      '</Relationships>';
  }

  function docRelsXml(ctx, hasHeader, hasFooter) {
    var rels = [
      ['rIdStyles', NS_DOC + '/styles', 'styles.xml'],
      ['rIdNumbering', NS_DOC + '/numbering', 'numbering.xml'],
      ['rIdSettings', NS_DOC + '/settings', 'settings.xml'],
      ['rIdFonts', NS_DOC + '/fontTable', 'fontTable.xml']
    ];
    var out = rels.map(function (r) {
      return '<Relationship Id="' + ctx.partIds[r[0]] + '" Type="' + r[1] + '" Target="' + r[2] + '"/>';
    });
    if (hasHeader) out.push('<Relationship Id="' + ctx.partIds.header + '" Type="' + NS_DOC + '/header" Target="header1.xml"/>');
    if (hasFooter) out.push('<Relationship Id="' + ctx.partIds.footer + '" Type="' + NS_DOC + '/footer" Target="footer1.xml"/>');
    ctx.rels.forEach(function (r) {
      out.push('<Relationship Id="' + r.id + '" Type="' + r.type + '" Target="' + U.escapeXml(r.target) + '"' +
        (r.mode ? ' TargetMode="' + r.mode + '"' : '') + '/>');
    });
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + CT_REL + '">' + out.join('') + '</Relationships>';
  }

  function corePropsXml(meta) {
    var now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    function t(tag, val) { return val ? '<' + tag + '>' + U.escapeXml(val) + '</' + tag + '>' : ''; }
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      t('dc:title', meta.title) + t('dc:subject', meta.subject) + t('dc:creator', meta.author) +
      t('cp:keywords', meta.keywords) + t('dc:description', meta.comments) +
      t('cp:lastModifiedBy', meta.lastModifiedBy || meta.author) + '<cp:revision>1</cp:revision>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + (meta.created || now) + '</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">' + now + '</dcterms:modified>' +
      '</cp:coreProperties>';
  }

  function appPropsXml(stats) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<Template>Normal.dotm</Template>' +
      '<Pages>' + (stats.pages || 1) + '</Pages>' +
      '<Words>' + (stats.words || 0) + '</Words>' +
      '<Characters>' + (stats.chars || 0) + '</Characters>' +
      '<Application>OpenWord</Application>' +
      '<DocSecurity>0</DocSecurity>' +
      '<AppVersion>1.0000</AppVersion>' +
      '</Properties>';
  }

  function settingsXml() {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:settings xmlns:w="' + NS_W + '">' +
      '<w:zoom w:percent="100"/>' +
      '<w:defaultTabStop w:val="720"/>' +
      '<w:characterSpacingControl w:val="compressPunctuation"/>' +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
      '</w:settings>';
  }

  function fontTableXml(fonts) {
    var names = Object.keys(fonts);
    if (!names.length) names = ['Calibri'];
    var xml = names.map(function (n) {
      return '<w:font w:name="' + U.escapeXml(n) + '"><w:family val="Auto"/><w:pitch val="default"/></w:font>';
    }).join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:fonts xmlns:w="' + NS_W + '">' + xml + '</w:fonts>';
  }

  /* ============================================================ main ===== */
  /**
   * @param {Object} opts
   *   editor   : element holding the document body
   *   header   : element holding header content (may be empty)
   *   footer   : element holding footer content (may be empty)
   *   page     : {size, orientation, margins{top,right,bottom,left,header,footer}}
   *   meta     : {title, author, subject, keywords, comments}
   *   stats    : {words, chars, pages}
   */
  function exportDocument(opts) {
    var ctx, page;
    return Promise.resolve().then(function () {
      var U2 = U;
      ctx = makeContext();
      page = Object.assign({
        size: 'Letter', orientation: 'portrait',
        margins: { top: 1, right: 1, bottom: 1, left: 1, header: 0.5, footer: 0.5 }
      }, opts.page || {});
      var size = PAGE_SIZES[page.size] || PAGE_SIZES.Letter;
      var wTwips = page.orientation === 'landscape' ? size[1] : size[0];
      ctx.page = {
        contentWidthTwips: wTwips - U2.inToTwip(page.margins.left) - U2.inToTwip(page.margins.right)
      };
      ctx.defaults = opts.defaults || {};

      // Pre-measure images so that even unloaded ones get sensible extents.
      IMG_SIZES = {};
      var imgs = opts.editor.querySelectorAll('img');
      return Promise.all(Array.prototype.map.call(imgs, function (img) {
        if (img.complete && img.naturalWidth) return Promise.resolve();
        return new Promise(function (res) {
          var done = false;
          function finish() { if (!done) { done = true; res(); } }
          img.addEventListener('load', finish, { once: true });
          img.addEventListener('error', finish, { once: true });
          setTimeout(finish, 1200);
          if (img.decode) img.decode().then(finish, finish);
        });
      })).then(function () {
        Array.prototype.forEach.call(imgs, function (img, i) {
          var key = img.getAttribute('src') || ('idx' + i);
          IMG_SIZES[key] = { w: img.naturalWidth || img.width, h: img.naturalHeight || img.height };
          img.dataset.wwKey = key;
        });
      });
    }).then(function () {
      var bodyItems = collectBlocks(opts.editor, ctx);
      var bodyXml = bodyItems.map(function (item) { return renderItem(item, ctx); }).join('');
      if (!/<w:p[ >]/.test(bodyXml)) bodyXml = '<w:p/>' + bodyXml;

      // Header / footer parts
      var hasHeader = hasContent(opts.header), hasFooter = hasContent(opts.footer);
      var headerXmlStr = '', footerXmlStr = '';
      if (hasHeader) headerXmlStr = hfXml(opts.header, ctx);
      if (hasFooter) footerXmlStr = hfXml(opts.footer, ctx);

      ctx.partIds = {
        styles: addRel(ctx, NS_DOC + '/styles', 'styles.xml'),
        numbering: addRel(ctx, NS_DOC + '/numbering', 'numbering.xml'),
        settings: addRel(ctx, NS_DOC + '/settings', 'settings.xml'),
        fonts: addRel(ctx, NS_DOC + '/fontTable', 'fontTable.xml')
      };
      if (hasHeader) ctx.partIds.header = addRel(ctx, NS_DOC + '/header', 'header1.xml');
      if (hasFooter) ctx.partIds.footer = addRel(ctx, NS_DOC + '/footer', 'footer1.xml');

      var documentXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        '<w:document xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '" ' +
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">' +
        '<w:body>' + bodyXml + sectionPropertiesXml(page, ctx,
          hasHeader ? [ctx.partIds.header] : [], hasFooter ? [ctx.partIds.footer] : []) +
        '</w:body></w:document>';

      var zip = new JSZip();
      zip.file('[Content_Types].xml', contentTypesXml(ctx, hasHeader, hasFooter));
      zip.folder('_rels').file('.rels', rootRelsXml());
      zip.folder('docProps')
        .file('core.xml', corePropsXml(opts.meta || {}))
        .file('app.xml', appPropsXml(opts.stats || {}));
      var wdir = zip.folder('word');
      wdir.file('document.xml', documentXml);
      wdir.file('styles.xml', stylesXml(ctx.defaults));
      wdir.file('numbering.xml', numberingXml(ctx));
      wdir.file('settings.xml', settingsXml());
      wdir.file('fontTable.xml', fontTableXml(collectFonts(opts)));
      wdir.folder('_rels').file('document.xml.rels', docRelsXml(ctx, hasHeader, hasFooter));
      if (hasHeader) wdir.file('header1.xml', headerXmlStr);
      if (hasFooter) wdir.file('footer1.xml', footerXmlStr);
      ctx.media.forEach(function (m) { wdir.folder('media').file(m.name.replace('media/', ''), m.bytes); });

      return zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      }).then(function (blob) {
        return { blob: blob, warnings: ctx.warnings, mediaCount: ctx.media.length };
      });
    });
  }

  function collectFonts(opts) {
    var fonts = { Calibri: 1, 'Calibri Light': 1, Arial: 1, Consolas: 1 };
    if (opts.defaults && opts.defaults.font) fonts[opts.defaults.font] = 1;
    try {
      var nodes = opts.editor.querySelectorAll('[style]');
      Array.prototype.forEach.call(nodes, function (el) {
        var fam = el.style.fontFamily;
        if (fam) fonts[fam.split(',')[0].replace(/["']/g, '').trim()] = 1;
      });
    } catch (e) { /* ignore */ }
    delete fonts[''];
    return fonts;
  }

  function hasContent(el) {
    if (!el) return false;
    if (U.stripZwsp(el.textContent || '').trim()) return true;
    if (el.querySelector('img, .ww-field, table')) return true;
    return false;
  }

  function hfXml(el, ctx) {
    var items = collectBlocks(el, ctx);
    var inner = items.map(function (item) { return renderItem(item, ctx); }).join('');
    if (!inner) inner = '<w:p/>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<w:hdr xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '" ' +
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' + inner + '</w:hdr>';
  }

  return {
    exportDocument: exportDocument,
    PAGE_SIZES: PAGE_SIZES,
    // exported for unit tests
    _internals: { pPrXml: pPrXml, rprXml: rprXml, collectBlocks: collectBlocks, columnWidths: columnWidths }
  };
})(WW.util);
