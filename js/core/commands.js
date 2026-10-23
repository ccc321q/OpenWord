/* ==========================================================================
   OpenWord — editing commands
   Everything the toolbar, menus, keyboard and context menu can do goes through
   WW.commands.exec(name, value) / queryState / queryValue.
   ========================================================================== */
WW.commands = (function (U, Sel) {
  'use strict';

  var editor = null;
  var state = {
    defaultFont: 'Calibri',
    defaultSize: 11,
    pendingFontSize: null,
    painter: null,          // {props:{...}} captured formatting
    lastTextColor: '#C00000',
    lastHighlight: '#FFFF00',
    imageHandles: null
  };

  function init() {
    editor = document.getElementById('editor');
  }

  function rootEl() {
    var r = Sel.rootOf(window.getSelection().anchorNode);
    return r || editor;
  }

  function ensureFocus() {
    var sel = window.getSelection();
    if (sel && sel.rangeCount && Sel.rootOf(sel.getRangeAt(0).startContainer)) return;
    var first = editor.firstElementChild || editor;
    editor.focus();
    Sel.caretAtStartOf(first);
  }

  /* ==================================================== block utilities == */
  var BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,pre,td,th,div';

  function blockFor(node, boundary) {
    if (!node) return null;
    if (node.nodeType === 3) node = node.parentElement;
    if (!node) return null;
    var stop = boundary || Sel.rootOf(node) || editor;
    var el = node;
    while (el && el !== stop) {
      var t = el.tagName;
      if (t === 'P' || t === 'LI' || /^H[1-6]$/.test(t) || t === 'BLOCKQUOTE' || t === 'PRE' ||
        t === 'TD' || t === 'TH' || t === 'DIV') return el;
      el = el.parentElement;
    }
    return null;
  }

  /** Blocks the selection touches, outermost-first, never crossing into table cells twice. */
  function selectedBlocks() {
    var range = Sel.liveRange();
    if (!range) return [];
    var root = Sel.rootOf(range.startContainer) || editor;
    var seen = [];
    var blocks = [];
    function add(node) {
      var b = blockFor(node, root);
      if (b && seen.indexOf(b) < 0) { seen.push(b); blocks.push(b); }
    }
    add(range.startContainer);
    add(range.endContainer);
    Sel.textNodesInRange(range).forEach(add);
    if (!blocks.length) {
      var cont = range.commonAncestorContainer;
      if (cont.nodeType === 1) {
        Array.prototype.forEach.call(cont.children, function (c) { add(c); });
      }
      if (!blocks.length) add(cont);
    }
    return blocks;
  }

  /** The block that contains the caret / start of the selection. */
  function anchorBlock() {
    var sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    return blockFor(sel.getRangeAt(0).startContainer);
  }

  function topBlockOf(el) {
    var node = el, top = el;
    while (node && node !== editor) {
      var p = node.parentElement;
      if (!p) break;
      if (p === editor || p.tagName === 'LI' || p.tagName === 'TD' || p.tagName === 'TH' || p.tagName === 'UL' || p.tagName === 'OL') break;
      node = p;
    }
    return node;
  }

  function newBlock(tag, src) {
    var el = document.createElement(tag);
    if (src) {
      var style = src.getAttribute('style');
      if (style) el.setAttribute('style', style);
      var cls = (src.getAttribute('class') || '').replace(/ww-list-paragraph|ww-pagebreak/g, '').trim();
      if (cls) el.setAttribute('class', cls);
      while (src.firstChild) el.appendChild(src.firstChild);
    }
    return el;
  }

  /* ===================================================== typed commands == */
  function execSimple(cmd) {
    document.execCommand('styleWithCSS', false, false);
    return document.execCommand(cmd, false, null);
  }

  function execStyled(cmd, value) {
    document.execCommand('styleWithCSS', false, true);
    var ok = document.execCommand(cmd, false, value);
    document.execCommand('styleWithCSS', false, false);
    return ok;
  }

  /** Replace the browser's placeholder font size with a real point size. */
  function applyFontSize(pt) {
    ensureFocus();
    var target = pt ? (Math.round(pt * 100) / 100) + 'pt' : state.defaultSize + 'pt';
    state.pendingFontSize = target;
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('fontSize', false, '7');
    document.execCommand('styleWithCSS', false, false);
    var root = rootEl();
    var nodes = root.querySelectorAll('[style]');
    var changed = 0;
    Array.prototype.forEach.call(nodes, function (el) {
      var fs = (el.style.fontSize || '').toLowerCase();
      if (fs === 'xxx-large' || fs === 'xx-large' && !el.dataset.wwKeepSize) {
        if (fs !== 'xxx-large') return;
        el.style.fontSize = target;
        changed++;
      }
    });
    cleanupEmptySpans(root);
    return true;
  }

  function applyFontFamily(family) {
    ensureFocus();
    var target = family || state.defaultFont;
    execStyled('fontName', target);
    var root = rootEl();
    Array.prototype.forEach.call(root.querySelectorAll('[style]'), function (el) {
      var fam = el.style.fontFamily;
      if (!fam) return;
      var first = fam.split(',')[0].replace(/["']/g, '').trim();
      if (first === target && el.style.length === 1 && el.tagName === 'SPAN') {
        el.style.fontFamily = target;    // normalise quoting
      }
    });
    cleanupEmptySpans(root);
    return true;
  }

  function applyColor(hex) {
    ensureFocus();
    state.lastTextColor = hex || '#000000';
    execStyled('foreColor', hex || '#000000');
    cleanupEmptySpans(rootEl());
    return true;
  }

  function applyHighlight(hex) {
    ensureFocus();
    var value = hex || 'transparent';
    state.lastHighlight = hex || null;
    document.execCommand('styleWithCSS', false, true);
    var ok = document.execCommand('hiliteColor', false, value);
    if (!ok) ok = document.execCommand('backColor', false, value);
    document.execCommand('styleWithCSS', false, false);
    // Highlight on a collapsed caret may be ignored: keep it as pending instead.
    if (Sel.isCollapsed()) pendingHighlights.push(value);
    cleanupEmptySpans(rootEl());
    return true;
  }
  var pendingHighlights = [];

  function cleanupEmptySpans(root) {
    Array.prototype.forEach.call(root.querySelectorAll('span'), function (sp) {
      if (sp.getAttribute('class') === 'ww-field') return;
      if (sp.getAttribute('class') === 'ww-caret-marker') return;
      var style = sp.getAttribute('style');
      if (style && /(?:^|;)\s*(font-family|font-size|color|background-color|text-decoration)\s*:\s*[^;]+/i.test(style.trim())) return;
      if (!style || !style.trim()) {
        if (sp.querySelector('img, .ww-field')) { unwrap(sp); return; }
        unwrap(sp);
      }
    });
    // Collapse redundant nesting: <b><b>x</b></b>
    Array.prototype.forEach.call(root.querySelectorAll('b>i>i, i>b>b'), function () { /* cheap no-op guard */ });
  }

  function unwrap(el) {
    var parent = el.parentNode;
    if (!parent) return;
    while (el.firstChild) parent.insertBefore(el.firstChild, el);
    parent.removeChild(el);
  }

  /* ========================================================= alignment === */
  var ALIGN = {
    justifyLeft: 'left', justifyCenter: 'center', justifyRight: 'right', justifyFull: 'justify'
  };

  function setAlignment(align) {
    var blocks = selectedBlocks();
    if (!blocks.length) return false;
    var current = alignOf(blocks[0]);
    var removing = current === align;
    blocks.forEach(function (b) {
      if (b.tagName === 'TD' || b.tagName === 'TH') {
        b.style.textAlign = align;
        Array.prototype.forEach.call(b.children, function (c) { if (c.nodeType === 1) c.style.textAlign = align; });
        return;
      }
      if (removing) b.style.removeProperty('text-align');
      else b.style.textAlign = align;
    });
    return true;
  }

  function alignOf(block) {
    var el = block;
    while (el && el !== editor) {
      if (el.style && el.style.textAlign) return el.style.textAlign.toLowerCase();
      el = el.parentElement;
    }
    var cs = window.getComputedStyle(block);
    return (cs.textAlign || 'left').toLowerCase() === 'start' ? 'left' : (cs.textAlign || 'left').toLowerCase();
  }

  /* ========================================================== spacing === */
  function setLineSpacing(value) {
    var blocks = selectedBlocks();
    if (!blocks.length) return false;
    blocks.forEach(function (b) {
      if (!value) b.style.removeProperty('line-height');
      else if (/^[\d.]+$/.test(String(value))) b.style.lineHeight = String(value);
      else b.style.lineHeight = value;
    });
    return true;
  }

  function setParagraphSpacing(opts) {
    var blocks = selectedBlocks();
    if (!blocks.length) return false;
    blocks.forEach(function (b) {
      if (opts.before !== undefined) b.style.marginTop = opts.before === null ? '' : opts.before + 'pt';
      if (opts.after !== undefined) b.style.marginBottom = opts.after === null ? '' : opts.after + 'pt';
      if (opts.left !== undefined) b.style.marginLeft = opts.left === null ? '' : opts.left + 'pt';
      if (opts.right !== undefined) b.style.marginRight = opts.right === null ? '' : opts.right + 'pt';
      if (opts.firstLine !== undefined) b.style.textIndent = opts.firstLine === null ? '' : opts.firstLine + 'pt';
      if (opts.align) b.style.textAlign = opts.align;
    });
    return true;
  }

  /* ============================================================ lists === */
  function listInfoOf(block) {
    var li = Sel.closest(block, 'li');
    if (!li) return null;
    var list = li.parentElement;
    if (!list || (list.tagName !== 'UL' && list.tagName !== 'OL')) return null;
    return { li: li, list: list, depth: listDepth(list) };
  }

  function listDepth(list) {
    var depth = 0, node = list;
    while (node && node !== editor) {
      if (node.tagName === 'UL' || node.tagName === 'OL') depth++;
      node = node.parentElement;
    }
    return depth;
  }

  function toggleList(tag) {
    ensureFocus();
    var blocks = selectedBlocks();
    if (!blocks.length) return false;
    var infos = blocks.map(listInfoOf);
    var allSame = infos.length > 0 && infos.every(function (i) { return i && i.list.tagName === tag.toUpperCase(); });
    if (allSame && tag === 'ul' && infos[0].list.tagName === 'UL') {
      infos.forEach(removeFromList);
      return true;
    }
    if (allSame) {
      // convert ul <-> ol
      infos.forEach(function (info) {
        var other = document.createElement(tag);
        copyStyleOf(info.list, other);
        if (tag === 'ul') other.removeAttribute('data-ww-list-type');
        other.innerHTML = info.list.innerHTML;
        info.list.replaceWith(other);
      });
      return true;
    }
    wrapBlocksInList(blocks, tag);
    return true;
  }

  function copyStyleOf(src, dst) {
    if (src.getAttribute('style')) dst.setAttribute('style', src.getAttribute('style'));
    if (src.getAttribute('data-ww-list-type')) dst.setAttribute('data-ww-list-type', src.getAttribute('data-ww-list-type'));
  }

  function wrapBlocksInList(blocks, tag) {
    var lastList = null;
    blocks.forEach(function (b) {
      if (b.tagName === 'LI') return;
      var holder = b.parentElement === editor ? b : topBlockOf(b);
      var info = listInfoOf(holder);
      if (info) {
        // move this block out of its list first, then re-wrap below
        var li = document.createElement('li');
        var p = newBlock('p', holder);
        li.appendChild(p);
        var nestedList = holder.parentElement && (holder.parentElement.tagName === 'UL' || holder.parentElement.tagName === 'OL') ? null : null;
        holder.replaceWith(li);
        holder = p;
      }
      var prev = holder.previousElementSibling;
      var target = lastList && prev === lastList ? lastList :
        (prev && prev.tagName === tag.toUpperCase() ? prev : null);
      if (!target) {
        target = document.createElement(tag);
        if (tag === 'ol') target.style.listStyleType = 'decimal';
        holder.parentElement.insertBefore(target, holder);
      }
      var li2 = document.createElement('li');
      var para = newBlock('p', holder);
      li2.appendChild(para);
      target.appendChild(li2);
      if (holder.parentNode) holder.remove();
      lastList = target;
      Sel.caretAtStartOf(para);
    });
  }

  function removeFromList(info) {
    var li = info.li, list = info.list;
    var parent = list.parentElement;
    var para = null;
    var kids = Array.prototype.slice.call(li.childNodes);
    var nestedLists = kids.filter(function (n) { return n.nodeType === 1 && (n.tagName === 'UL' || n.tagName === 'OL'); });
    var inlineHolder = document.createElement('p');
    kids.forEach(function (n) {
      if (nestedLists.indexOf(n) >= 0) return;
      if (n.nodeType === 1 && /^(P|H[1-6]|DIV|BLOCKQUOTE|PRE|TABLE)$/.test(n.tagName)) {
        parent.insertBefore(n, list);
        return;
      }
      inlineHolder.appendChild(n);
    });
    if (inlineHolder.childNodes.length || !para) {
      if (!inlineHolder.childNodes.length) inlineHolder.appendChild(document.createElement('br'));
      parent.insertBefore(inlineHolder, list);
      para = inlineHolder;
    }
    li.remove();
    // Re-attach nested lists at this level
    nestedLists.forEach(function (nl) { parent.insertBefore(nl, list); });
    if (!list.children.length) list.remove();
    if (para) Sel.caretAtStartOf(para);
  }

  function changeListIndent(dir) {
    ensureFocus();
    var blocks = selectedBlocks();
    var infos = blocks.map(listInfoOf).filter(Boolean);
    if (infos.length) {
      if (dir > 0) {
        infos.slice().reverse().forEach(function (info) {
          var prev = info.li.previousElementSibling;
          if (!prev || prev.tagName !== 'LI') return;
          var nested = prev.querySelector(':scope > ul, :scope > ol');
          if (!nested) {
            nested = document.createElement(info.list.tagName);
            prev.appendChild(nested);
          }
          nested.appendChild(info.li);
        });
      } else {
        infos.forEach(function (info) {
          var outer = Sel.closest(info.list.parentElement, 'li');
          if (!outer) {
            // promote to plain paragraph
            removeFromList(info);
            return;
          }
          var parentList = info.list.parentElement === outer ? info.list : null;
          var insertAfter = outer;
          var kids2 = Array.prototype.slice.call(info.list.children);
          var idx = kids2.indexOf(info.li);
          var before = kids2.slice(0, idx), after = kids2.slice(idx + 1);
          info.list.remove();
          before.forEach(function (b) {
            var l = document.createElement('ul');
            outer.parentElement.insertBefore(l, outer);
            l.appendChild(b);
            l.dataset.wwNumId = info.list.dataset.wwNumId || '';
          });
          outer.parentElement.insertBefore(info.li, insertAfter.nextSibling);
          after.forEach(function (b) {
            var l = document.createElement(info.list.tagName);
            info.li.appendChild(l);
            l.appendChild(b);
          });
        });
      }
      return true;
    }
    // plain indentation
    var step = 18; // pt
    blocks.forEach(function (b) {
      if (b.tagName === 'LI') return;
      var cur = U.cssLengthToPt(b.style.marginLeft, b) || 0;
      var next = Math.max(0, cur + dir * step);
      b.style.marginLeft = next ? next + 'pt' : '';
    });
    return true;
  }

  /* ======================================================= block styles == */
  var STYLE_MAP = {
    'Normal': { tag: 'p', cls: '' },
    'p': { tag: 'p', cls: '' },
    'Heading1': { tag: 'h1' }, 'Heading2': { tag: 'h2' }, 'Heading3': { tag: 'h3' },
    'Heading4': { tag: 'h4' }, 'Heading5': { tag: 'h5' }, 'Heading6': { tag: 'h6' },
    'Title': { tag: 'p', cls: 'ww-title' },
    'Subtitle': { tag: 'p', cls: 'ww-subtitle' },
    'Quote': { tag: 'p', cls: 'ww-quote' },
    'IntenseQuote': { tag: 'p', cls: 'ww-intense-quote' },
    'ListParagraph': { tag: 'p', cls: '' },
    'Code': { tag: 'pre' }
  };

  function applyBlockStyle(styleId) {
    ensureFocus();
    var spec = STYLE_MAP[styleId] || STYLE_MAP.Normal;
    var blocks = selectedBlocks();
    if (!blocks.length) return false;
    blocks.forEach(function (b) {
      if (b.tagName === 'TD' || b.tagName === 'TH' || b.tagName === 'LI') {
        // operate on the first paragraph inside instead of the cell itself
        var inner = b.querySelector('p,h1,h2,h3,h4,h5,h6');
        if (inner) b = inner;
      }
      var el = newBlock(spec.tag === 'pre' ? 'pre' : spec.tag, b);
      var cls = (b.getAttribute('class') || '').replace(/ww-title|ww-subtitle|ww-quote|ww-intense-quote|ww-list-paragraph/g, '').trim();
      if (spec.cls) el.setAttribute('class', (cls + ' ' + spec.cls).trim());
      else if (cls) el.setAttribute('class', cls); else el.removeAttribute('class');
      b.replaceWith(el);
    });
    return true;
  }

  function currentBlockStyle() {
    var b = anchorBlock();
    if (!b) return 'Normal';
    var t = b.tagName.toLowerCase();
    if (/^h[1-6]$/.test(t)) return 'Heading' + t.charAt(1);
    if (t === 'pre') return 'Code';
    if (t === 'blockquote') return 'Quote';
    var cls = b.getAttribute('class') || '';
    if (/ww-title/.test(cls)) return 'Title';
    if (/ww-subtitle/.test(cls)) return 'Subtitle';
    if (/ww-intense-quote/.test(cls)) return 'IntenseQuote';
    if (/ww-quote/.test(cls)) return 'Quote';
    return 'Normal';
  }

  /* ======================================================== clear style == */
  function clearFormatting() {
    ensureFocus();
    document.execCommand('styleWithCSS', false, true);
    document.execCommand('removeFormat', false, null);
    document.execCommand('styleWithCSS', false, false);
    var blocks = selectedBlocks();
    blocks.forEach(function (b) {
      if (b.tagName === 'TD' || b.tagName === 'TH') return;
      var keep = { marginLeft: b.style.marginLeft, marginRight: b.style.marginRight };
      b.style.removeProperty('text-align');
      b.style.removeProperty('line-height');
      b.style.removeProperty('text-indent');
      b.style.removeProperty('background-color');
      b.style.removeProperty('font-size');
      b.style.removeProperty('font-family');
      b.style.removeProperty('color');
      b.style.removeProperty('margin-top');
      b.style.removeProperty('margin-bottom');
      if (/^H[1-6]$/.test(b.tagName) || b.tagName === 'BLOCKQUOTE' || b.tagName === 'PRE') {
        var p = newBlock('p', b);
        b.replaceWith(p);
      } else {
        b.setAttribute('class', (b.getAttribute('class') || '').replace(/ww-title|ww-subtitle|ww-quote|ww-intense-quote/g, '').trim());
      }
    });
    return true;
  }

  /* ============================================================ insert == */
  function insertHTML(html, selectAfter) {
    ensureFocus();
    document.execCommand('styleWithCSS', false, false);
    document.execCommand('insertHTML', false, html);
    return true;
  }

  function insertText(text) {
    ensureFocus();
    document.execCommand('insertText', false, text);
    return true;
  }

  function insertLink(url, text) {
    ensureFocus();
    var sel = window.getSelection();
    var selected = sel && !sel.isCollapsed ? sel.toString() : '';
    var content = text || selected || url;
    if (selected) {
      document.execCommand('styleWithCSS', false, true);
      document.execCommand('createLink', false, url);
      document.execCommand('styleWithCSS', false, false);
    } else {
      insertHTML('<a href="' + U.escapeHtml(url) + '">' + U.escapeHtml(content) + '</a>&nbsp;');
    }
    return true;
  }

  function insertImageDataUrl(dataUrl, opts) {
    opts = opts || {};
    ensureFocus();
    var maxW = opts.width || 0;
    var img = new Image();
    img.onload = function () {
      var w = img.naturalWidth, h = img.naturalHeight;
      var limit = opts.maxWidth || 624;
      if (w > limit) { h = Math.round(h * limit / w); w = limit; }
      if (maxW) { h = Math.round(h * maxW / w); w = maxW; }
      var align = opts.align === 'center' ? 'display:block;margin:8pt auto;' :
        (opts.align === 'right' ? 'float:right;margin:0 0 8pt 12px;' : (opts.align === 'left' ? 'float:left;margin:0 12px 8pt 0;' : ''));
      var alt = U.escapeHtml(opts.alt || '');
      insertHTML('<img src="' + dataUrl + '" alt="' + alt + '" style="width:' + w + 'px;height:' + Math.round(h) + 'px;' + align + '">');
      WW.events && WW.events.changed('insert-image');
    };
    img.onerror = function () { WW.events && WW.events.toast && WW.events.toast('That image could not be loaded', 'error'); };
    img.src = dataUrl;
    return true;
  }

  function insertTable(rows, cols, opts) {
    opts = opts || {};
    ensureFocus();
    var contentW = (WW.pages && WW.pages.contentWidthPx) ? WW.pages.contentWidthPx() : 624;
    var cellW = Math.floor(contentW / cols);
    var html = '<table style="width:' + contentW + 'px;border-collapse:collapse">';
    if (cols) {
      html += '<colgroup>';
      for (var c = 0; c < cols; c++) html += '<col width="' + cellW + '">';
      html += '</colgroup>';
    }
    for (var r = 0; r < rows; r++) {
      html += '<tr>';
      for (var c2 = 0; c2 < cols; c2++) {
        var isHead = r === 0 && (opts.header || opts.headerRow);
        html += '<' + (isHead ? 'th' : 'td') + ' style="border:1px solid #ababab;padding:4pt 6pt;width:' + cellW + 'px">' +
          '<p></p></' + (isHead ? 'th' : 'td') + '>';
      }
      html += '</tr>';
    }
    html += '</table><p><br></p>';
    insertHTML(html);
    // place caret in the first cell
    var sel = window.getSelection();
    var node = sel.anchorNode;
    var tbl = null;
    var el = node && node.nodeType === 1 ? node : (node && node.parentElement);
    if (el) tbl = el.closest('table') || (el.parentElement && el.parentElement.querySelector('table'));
    if (!tbl) {
      var all = editor.querySelectorAll('table');
      if (all.length) tbl = all[all.length - 1];
    }
    if (tbl && tbl.rows[0] && tbl.rows[0].cells[0]) {
      Sel.caretAtStartOf(tbl.rows[0].cells[0].firstElementChild || tbl.rows[0].cells[0]);
    }
    return true;
  }

  /* ------------------------------------------------------ table editing */
  function tableContext() {
    var sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    var el = sel.anchorNode && (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement);
    if (!el) return null;
    var cell = el.closest('td,th');
    if (!cell) return null;
    var row = cell.parentElement;
    var table = row.closest('table');
    if (!table || !editor.contains(table)) return null;
    return { table: table, row: row, cell: cell, rowIndex: Array.prototype.indexOf.call(table.rows, row), cellIndex: Array.prototype.indexOf.call(row.cells, cell) };
  }

  function insertRow(ctxCtx, above) {
    var c = ctxCtx || tableContext(); if (!c) return false;
    var cols = c.table.rows[0].cells.length;
    var tr = document.createElement('tr');
    for (var i = 0; i < cols; i++) {
      var td = document.createElement('td');
      var p = document.createElement('p');
      td.appendChild(p);
      tr.appendChild(td);
    }
    c.row.parentElement.insertBefore(tr, above ? c.row : c.row.nextSibling);
    inheritCellStyles(c.row, tr);
    return true;
  }

  function inheritCellStyles(fromRow, toRow) {
    for (var i = 0; i < toRow.cells.length && i < fromRow.cells.length; i++) {
      toRow.cells[i].setAttribute('style', fromRow.cells[i].getAttribute('style') || '');
      if (fromRow.cells[i].tagName === 'TH') {
        var th = document.createElement('th');
        th.setAttribute('style', fromRow.cells[i].getAttribute('style') || '');
        th.appendChild(document.createElement('p'));
        toRow.cells[i].replaceWith(th);
      }
    }
  }

  function insertColumn(ctxCtx, before) {
    var c = ctxCtx || tableContext(); if (!c) return false;
    var idx = before ? c.cellIndex : c.cellIndex + 1;
    Array.prototype.forEach.call(c.table.rows, function (row, ri) {
      var ref = row.cells[Math.min(idx, row.cells.length - 1)];
      var td = document.createElement(ri === 0 && ref && ref.tagName === 'TH' ? 'th' : 'td');
      td.appendChild(document.createElement('p'));
      if (ref) { td.setAttribute('style', ref.getAttribute('style') || ''); row.insertBefore(td, before ? ref : ref.nextSibling); }
      else row.appendChild(td);
    });
    var cg = c.table.querySelector('colgroup');
    if (cg) {
      var col = document.createElement('col');
      var refCol = cg.children[Math.min(idx, cg.children.length - 1)];
      if (refCol) { col.setAttribute('width', refCol.getAttribute('width') || ''); cg.insertBefore(col, before ? refCol : refCol.nextSibling); }
      else cg.appendChild(col);
    }
    resizeTable(c.table);
    return true;
  }

  function deleteRow(ctxCtx) {
    var c = ctxCtx || tableContext(); if (!c) return false;
    if (c.table.rows.length <= 1) { deleteTable(c); return true; }
    c.row.remove();
    return true;
  }

  function deleteColumn(ctxCtx) {
    var c = ctxCtx || tableContext(); if (!c) return false;
    var idx = c.cellIndex;
    var only = c.table.rows[0].cells.length <= 1;
    if (only) { deleteTable(c); return true; }
    Array.prototype.forEach.call(c.table.rows, function (row) {
      if (row.cells[idx]) row.cells[idx].remove();
    });
    var cg = c.table.querySelector('colgroup');
    if (cg && cg.children[idx]) cg.children[idx].remove();
    resizeTable(c.table);
    return true;
  }

  function deleteTable(ctxCtx) {
    var c = ctxCtx || tableContext(); if (!c) return false;
    var p = document.createElement('p');
    c.table.replaceWith(p);
    Sel.caretAtStartOf(p);
    return true;
  }

  function mergeCells() {
    var range = Sel.liveRange();
    var c = tableContext();
    if (!c) return false;
    var cells = [];
    if (range) {
      var tds = range.commonAncestorContainer.querySelectorAll ?
        Array.prototype.filter.call(range.commonAncestorContainer.querySelectorAll('td,th'), function (td) { return range.intersectsNode(td) || td === c.cell; }) : [];
      if (tds.length) cells = tds;
    }
    if (cells.length < 2) return false;
    var rows = cells.map(function (td) { return td.parentElement; });
    var first = cells[0];
    var rowSpan = new Set(rows).size;
    var colSpan = Math.max.apply(null, Array.from(new Set(rows)).map(function (r) {
      return cells.filter(function (td) { return td.parentElement === r; }).length;
    }));
    if (new Set(rows).size !== rows.length && rows.length > 1) {
      // Non-rectangular selection: merge the row span only if all cells share a column set
      colSpan = cells.filter(function (td) { return td.parentElement === rows[0]; }).length;
      rowSpan = rows.length;
    }
    var text = cells.map(function (td) { return td.textContent; }).filter(Boolean).join('\n');
    var lines = text.split('\n').filter(function (l) { return l.trim(); });
    first.innerHTML = '';
    if (!lines.length) first.appendChild(document.createElement('p'));
    lines.forEach(function (line) {
      var p = document.createElement('p');
      p.textContent = line;
      first.appendChild(p);
    });
    first.colSpan = colSpan;
    first.rowSpan = rowSpan;
    cells.slice(1).forEach(function (td) {
      var row = td.parentElement;
      td.remove();
      if (!row.cells.length) row.remove();
    });
    padTable(c.table);
    return true;
  }

  function splitCell() {
    var c = tableContext();
    if (!c) return false;
    var cell = c.cell;
    if ((cell.colSpan || 1) <= 1 && (cell.rowSpan || 1) <= 1) return false;
    var cs = cell.colSpan || 1, rs = cell.rowSpan || 1;
    cell.colSpan = 1;
    cell.rowSpan = 1;
    var row = cell.parentElement;
    for (var i = 1; i < cs; i++) {
      var td = document.createElement(cell.tagName.toLowerCase());
      td.appendChild(document.createElement('p'));
      td.setAttribute('style', cell.getAttribute('style') || '');
      row.insertBefore(td, cell.nextSibling);
    }
    for (var r = 1; r < rs; r++) {
      var target = row.nextElementSibling;
      if (!target) break;
      var td2 = document.createElement(cell.tagName.toLowerCase());
      td2.appendChild(document.createElement('p'));
      td2.setAttribute('style', cell.getAttribute('style') || '');
      target.insertBefore(td2, target.cells[Math.min(c.cellIndex + cs, target.cells.length)] || null);
    }
    padTable(c.table);
    resizeTable(c.table);
    return true;
  }

  function padTable(table) {
    var max = 0;
    Array.prototype.forEach.call(table.rows, function (r) {
      var n = 0;
      Array.prototype.forEach.call(r.cells, function (c) { n += (c.colSpan || 1); });
      max = Math.max(max, n);
    });
    Array.prototype.forEach.call(table.rows, function (r) {
      var n = 0;
      Array.prototype.forEach.call(r.cells, function (c) { n += (c.colSpan || 1); });
      while (n < max) {
        var td = document.createElement(r.cells[0] ? r.cells[0].tagName.toLowerCase() : 'td');
        td.appendChild(document.createElement('p'));
        td.style.border = '1px solid #ababab';
        td.style.padding = '4pt 6pt';
        r.appendChild(td);
        n++;
      }
    });
  }

  function resizeTable(table) {
    var cols = table.rows[0] ? table.rows[0].cells.length : 0;
    if (!cols) return;
    var contentW = (WW.pages && WW.pages.contentWidthPx) ? WW.pages.contentWidthPx() : 624;
    var w = Math.floor(contentW / cols);
    table.style.width = contentW + 'px';
    var cg = table.querySelector('colgroup');
    if (cg) {
      while (cg.children.length < table.rows[0].cells.length) {
        var col = document.createElement('col'); cg.appendChild(col);
      }
      while (cg.children.length > table.rows[0].cells.length) cg.lastChild.remove();
      Array.prototype.forEach.call(cg.children, function (c) { c.setAttribute('width', w); });
    }
  }

  function setCellShading(hex) {
    var c = tableContext(); if (!c) return false;
    Array.prototype.forEach.call(c.row.cells, function () { });
    var target = c.cell;
    if (hex) target.style.backgroundColor = hex; else target.style.removeProperty('background-color');
    return true;
  }

  function setCellVerticalAlign(val) {
    var c = tableContext(); if (!c) return false;
    c.cell.style.verticalAlign = val;
    return true;
  }

  function toggleHeaderRow() {
    var c = tableContext(); if (!c) return false;
    var row = c.table.rows[0];
    var isHeader = row.cells[0] && row.cells[0].tagName === 'TH';
    Array.prototype.forEach.call(row.cells, function (cell) {
      var rep = document.createElement(isHeader ? 'td' : 'th');
      rep.setAttribute('style', cell.getAttribute('style') || '');
      rep.innerHTML = cell.innerHTML;
      rep.colSpan = cell.colSpan; rep.rowSpan = cell.rowSpan;
      cell.replaceWith(rep);
    });
    c.table.setAttribute('data-header', isHeader ? 'false' : 'true');
    return true;
  }

  /* -------------------------------------------------------- page parts -- */
  function insertPageBreak() {
    ensureFocus();
    var block = anchorBlock();
    var el = document.createElement('div');
    el.className = 'ww-pagebreak';
    el.appendChild(document.createElement('p'));
    if (block && block.parentElement) {
      block.parentElement.insertBefore(el, block.nextSibling);
      var p = el.firstElementChild;
      Sel.caretAtStartOf(p);
    } else {
      editor.appendChild(el);
      Sel.caretAtStartOf(el.firstElementChild);
    }
    return true;
  }

  function removePageBreak() {
    var block = anchorBlock();
    var wrap = block && block.parentElement;
    if (wrap && wrap.classList && wrap.classList.contains('ww-pagebreak')) {
      var p = document.createElement('p');
      wrap.replaceWith(p);
      Sel.caretAtStartOf(p);
      return true;
    }
    return false;
  }

  function fieldHtml(kind) {
    var text = kind === 'page' ? '1' : kind === 'numpages' ? '1' : kind === 'date' ?
      new Date().toLocaleDateString() : kind === 'time' ? new Date().toLocaleTimeString() : '';
    return '<span class="ww-field" data-ww-field="' + kind + '" contenteditable="false">' + U.escapeHtml(text) + '</span>';
  }

  function insertPageNumber() {
    ensureFocus();
    insertHTML(fieldHtml('page') + ' ' + 'of ' + fieldHtml('numpages') + ' ');
    return true;
  }

  function insertDate(withTime) {
    ensureFocus();
    var d = new Date();
    var text = d.toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' });
    if (withTime) text += ' ' + d.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' });
    return insertText(text);
  }

  function insertToc() {
    var headings = U.$$('h1,h2,h3', editor);
    if (!headings.length) { WW.events && WW.events.toast && WW.events.toast('Add some headings first', 'warn'); return false; }
    var html = '<p class="ww-title" style="font-size:16pt">Contents</p><ul style="list-style:none;padding-left:0">';
    headings.forEach(function (h, i) {
      var level = parseInt(h.tagName.charAt(1), 10);
      var indent = (level - 1) * 14;
      html += '<li style="margin-left:' + indent + 'pt">' + U.escapeHtml(h.textContent.trim() || '(untitled)') + '</li>';
    });
    html += '</ul><p><br></p>';
    var block = anchorBlock();
    var el = document.createElement('div');
    el.innerHTML = html;
    if (block && block.parentElement) block.parentElement.insertBefore(el, block.nextSibling);
    else editor.appendChild(el);
    WW.events && WW.events.changed('insert-toc');
    return true;
  }

  /* =================================================== format painter == */
  function capturePainter() {
    var rp = caretInlineProps();
    state.painter = rp;
    document.body.classList.add('painter-active');
    return true;
  }
  function applyPainter() {
    if (!state.painter) return false;
    var p = state.painter;
    ensureFocus();
    if (p.bold !== null) execSimple(p.bold ? 'bold' : 'bold');
    document.execCommand('styleWithCSS', false, true);
    if (p.color) document.execCommand('foreColor', false, p.color);
    if (p.highlight) document.execCommand('hiliteColor', false, p.highlight);
    if (p.font) document.execCommand('fontName', false, p.font);
    document.execCommand('styleWithCSS', false, false);
    if (p.size) applyFontSize(p.size);
    return true;
  }
  function releasePainter() {
    state.painter = null;
    document.body.classList.remove('painter-active');
  }

  /* ============================================================ query === */
  function caretNode() {
    var sel = window.getSelection();
    if (!sel || !sel.rangeCount) return null;
    var r = sel.getRangeAt(0);
    var node = r.startContainer;
    return node.nodeType === 3 ? node.parentElement : node;
  }

  function caretInlineProps() {
    var el = caretNode();
    if (!el) return {};
    var cs = window.getComputedStyle(el);
    var props = {
      bold: parseInt(cs.fontWeight, 10) >= 600,
      italic: cs.fontStyle === 'italic',
      underline: (cs.textDecorationLine || '').indexOf('underline') >= 0,
      strike: (cs.textDecorationLine || '').indexOf('line-through') >= 0,
      color: U.rgbStringToHex(cs.color),
      highlight: (function () {
        var bg = U.rgbStringToHex(cs.backgroundColor);
        return bg && bg !== '#FFFFFF' && cs.backgroundColor.indexOf('rgba(0, 0, 0, 0)') < 0 ? bg : null;
      })(),
      font: (cs.fontFamily || '').split(',')[0].replace(/["']/g, '').trim(),
      size: Math.round(U.pxToPt(parseFloat(cs.fontSize)) * 10) / 10,
      vertAlign: cs.verticalAlign
    };
    return props;
  }

  function queryState(cmd) {
    var el = caretNode();
    try {
      switch (cmd) {
        case 'bold': return hasTag('b,strong') || parseInt(window.getComputedStyle(el).fontWeight, 10) >= 700;
        case 'italic': return hasTag('i,em') || window.getComputedStyle(el).fontStyle === 'italic';
        case 'underline': return hasTag('u') || (window.getComputedStyle(el).textDecorationLine || '').indexOf('underline') >= 0;
        case 'strikeThrough': return hasTag('s,strike,del') || (window.getComputedStyle(el).textDecorationLine || '').indexOf('line-through') >= 0;
        case 'subscript': return hasTag('sub') || window.getComputedStyle(el).verticalAlign === 'sub';
        case 'superscript': return hasTag('sup') || window.getComputedStyle(el).verticalAlign === 'super';
        case 'insertUnorderedList': return !!Sel.closest(el, 'ul');
        case 'insertOrderedList': return !!Sel.closest(el, 'ol');
        case 'justifyLeft': return alignOf(anchorBlock() || editor) === 'left';
        case 'justifyCenter': return alignOf(anchorBlock() || editor) === 'center';
        case 'justifyRight': return alignOf(anchorBlock() || editor) === 'right';
        case 'justifyFull': return alignOf(anchorBlock() || editor) === 'justify';
        case 'paste': case 'cut': case 'copy': return false;
        default: return false;
      }
    } catch (e) { return false; }
  }

  function hasTag(selStr) {
    var el = caretNode();
    if (!el) return false;
    var node = el;
    var root = Sel.rootOf(el);
    while (node && node !== root) {
      if (node.matches && node.matches(selStr)) return true;
      node = node.parentElement;
    }
    return false;
  }

  function queryValue(cmd) {
    var el = caretNode();
    if (!el) return '';
    var cs = window.getComputedStyle(el);
    switch (cmd) {
      case 'font-family': return (cs.fontFamily || '').split(',')[0].replace(/["']/g, '').trim();
      case 'font-size': return String(Math.round(U.pxToPt(parseFloat(cs.fontSize)) * 10) / 10);
      case 'foreColor': return U.rgbStringToHex(cs.color) || '#000000';
      case 'hiliteColor': {
        var bg = U.rgbStringToHex(cs.backgroundColor);
        return (bg && cs.backgroundColor.indexOf('rgba(0, 0, 0, 0)') < 0) ? bg : '';
      }
      case 'block-style': return currentBlockStyle();
      case 'line-spacing': {
        var b = anchorBlock();
        if (!b) return '';
        var lh = b.style.lineHeight || '';
        return lh;
      }
      default: return '';
    }
  }

  /* ==================================================== caret helpers === */
  function removeEmptyBlocksAroundCaret() {
    // After removeFormat / list removal we can end up with <p><br></p><p><br></p>
    var b = anchorBlock();
    if (!b) return;
    var next = b.nextElementSibling;
    while (next && !next.textContent.trim() && !next.querySelector('img,table,.ww-field') &&
      b.textContent.trim() && !b.querySelector('img,table,.ww-field')) {
      var n2 = next.nextElementSibling;
      next.remove();
      next = n2;
    }
  }

  /* ========================================================== dispatch == */
  function exec(cmd, value) {
    if (!editor) init();
    switch (cmd) {
      case 'bold': case 'italic': case 'underline': case 'strikeThrough':
      case 'subscript': case 'superscript':
        return execSimple(cmd);
      case 'font-family': return applyFontFamily(value);
      case 'font-size': return applyFontSize(value ? parseFloat(value) : null);
      case 'increase-size': {
        var sizes = WW.APP && WW.APP.FONT_SIZES ? WW.APP.FONT_SIZES : [8, 9, 10, 11, 12, 14, 18, 24, 36, 48, 72];
        var cur = parseFloat(queryValue('font-size'));
        var next = sizes.find(function (s) { return s > cur; }) || sizes[sizes.length - 1];
        return applyFontSize(next);
      }
      case 'decrease-size': {
        var sizes2 = WW.APP && WW.APP.FONT_SIZES ? WW.APP.FONT_SIZES : [8, 9, 10, 11, 12, 14, 18, 24, 36, 48, 72];
        var cur2 = parseFloat(queryValue('font-size'));
        var prev = sizes2.slice().reverse().find(function (s) { return s < cur2; }) || sizes2[0];
        return applyFontSize(prev);
      }
      case 'foreColor': return applyColor(value || state.lastTextColor);
      case 'hiliteColor': return applyHighlight(value || state.lastHighlight);
      case 'clear-formatting': return clearFormatting();
      case 'justifyLeft': case 'justifyCenter': case 'justifyRight': case 'justifyFull':
        return setAlignment(ALIGN[cmd]);
      case 'setAlignment': return setAlignment(value);
      case 'insertUnorderedList': return toggleList('ul');
      case 'insertOrderedList': return toggleList('ol');
      case 'indent': return changeListIndent(1);
      case 'outdent': return changeListIndent(-1);
      case 'formatBlock': return applyBlockStyle(value);
      case 'setLineSpacing': return setLineSpacing(value);
      case 'setParagraphSpacing': return setParagraphSpacing(value || {});
      case 'insertHorizontalRule': return insertHTML('<hr><p><br></p>');
      case 'insert-page-break': return insertPageBreak();
      case 'remove-page-break': return removePageBreak();
      case 'insert-page-number': return insertPageNumber();
      case 'insert-date': return insertDate(false);
      case 'insert-datetime': return insertDate(true);
      case 'insert-toc': return insertToc();
      case 'insert-table': return insertTable(value.rows, value.cols, value);
      case 'insert-link': return insertLink(value.url, value.text);
      case 'insert-image': return insertImageDataUrl(value.dataUrl, value);
      case 'remove-link': return removeLink();
      case 'select-all': ensureFocus(); document.execCommand('selectAll'); return true;
      case 'format-painter': return state.painter ? (applyPainter(), true) : capturePainter();
      case 'table-insert-row-above': return insertRow(null, true);
      case 'table-insert-row-below': return insertRow(null, false);
      case 'table-insert-col-before': return insertColumn(null, true);
      case 'table-insert-col-after': return insertColumn(null, false);
      case 'table-delete-row': return deleteRow();
      case 'table-delete-col': return deleteColumn();
      case 'table-delete': return deleteTable();
      case 'table-merge': return mergeCells();
      case 'table-split': return splitCell();
      case 'table-header-row': return toggleHeaderRow();
      case 'table-cell-shading': return setCellShading(value);
      case 'table-cell-valign': return setCellVerticalAlign(value);
      case 'paste-plain': return pastePlainText(value);
      case 'insert-text': return insertText(value);
      case 'insert-html': return insertHTML(value);
      default: return false;
    }
  }

  function removeLink() {
    ensureFocus();
    var el = caretNode();
    var a = Sel.closest(el, 'a');
    if (!a) return false;
    unwrap(a);
    return true;
  }

  function pastePlainText(text) {
    ensureFocus();
    if (text == null) {
      if (navigator.clipboard && navigator.clipboard.readText) {
        return navigator.clipboard.readText().then(function (t) {
          insertText(t);
          WW.events && WW.events.changed('paste');
          return true;
        }).catch(function () {
          WW.events && WW.events.toast && WW.events.toast('Press Ctrl+Shift+V to paste as plain text', 'info');
          return false;
        });
      }
      return false;
    }
    insertText(text);
    WW.events && WW.events.changed('paste');
    return true;
  }

  /* ------------------------------------------------ image selection UI - */
  function attachImageUI() {
    editor.addEventListener('click', function (e) {
      var img = e.target && e.target.tagName === 'IMG' ? e.target : null;
      if (img) selectImage(img);
      else clearImageSelection();
    });
    document.addEventListener('mousedown', function (e) {
      if (!e.target.closest || !e.target.closest('.ww-img-resizer')) clearImageSelection();
    }, true);
  }

  function selectImage(img) {
    clearImageSelection();
    img.classList.add('ww-selected');
    var box = document.createElement('div');
    box.className = 'ww-img-resizer';
    var handle = document.createElement('div');
    handle.className = 'ww-img-handle se';
    box.appendChild(handle);
    document.body.appendChild(box);
    state.imageHandles = { img: img, box: box };
    positionImageBox();
    var startX = 0, startY = 0, startW = 0, startH = 0;
    handle.addEventListener('mousedown', function (e) {
      e.preventDefault();
      startX = e.clientX; startY = e.clientY;
      startW = img.getBoundingClientRect().width; startH = img.getBoundingClientRect().height;
      var zoom = (WW.APP && WW.APP.zoom) || 1;
      function onMove(ev) {
        var dw = (ev.clientX - startX) / zoom;
        var w = Math.max(24, Math.round(startW + dw));
        var h = Math.round(startH * (w / startW));
        img.style.width = w + 'px';
        img.style.height = h + 'px';
        positionImageBox();
      }
      function onUp() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        WW.history && WW.history.push('resize image');
        WW.events && WW.events.changed('resize-image');
      }
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
  }

  function positionImageBox() {
    if (!state.imageHandles) return;
    var img = state.imageHandles.img, box = state.imageHandles.box;
    var r = img.getBoundingClientRect();
    box.style.left = (r.left + window.scrollX) + 'px';
    box.style.top = (r.top + window.scrollY) + 'px';
    box.style.width = r.width + 'px';
    box.style.height = r.height + 'px';
  }

  function clearImageSelection() {
    if (!state.imageHandles) return;
    state.imageHandles.img.classList.remove('ww-selected');
    state.imageHandles.box.remove();
    state.imageHandles = null;
  }

  /* =========================================================== exports == */
  var API = {
    init: init, exec: exec, queryState: queryState, queryValue: queryValue,
    blockFor: blockFor, selectedBlocks: selectedBlocks, anchorBlock: anchorBlock,
    applyBlockStyle: applyBlockStyle, currentBlockStyle: currentBlockStyle,
    setAlignment: setAlignment, alignOf: alignOf, setLineSpacing: setLineSpacing,
    setParagraphSpacing: setParagraphSpacing, toggleList: toggleList, changeListIndent: changeListIndent,
    clearFormatting: clearFormatting, insertHTML: insertHTML, insertText: insertText,
    insertLink: insertLink, insertImageDataUrl: insertImageDataUrl, insertTable: insertTable,
    insertPageBreak: insertPageBreak, insertPageNumber: insertPageNumber, insertToc: insertToc,
    removeLink: removeLink, tableContext: tableContext, attachImageUI: attachImageUI,
    selectImage: selectImage, clearImageSelection: clearImageSelection, positionImageBox: positionImageBox,
    caretInlineProps: caretInlineProps, caretNode: caretNode, releasePainter: releasePainter,
    applyPainter: applyPainter, capturePainter: capturePainter, cleanupEmptySpans: cleanupEmptySpans,
    listInfoOf: listInfoOf, padTable: padTable, resizeTable: resizeTable,
    fieldHtml: fieldHtml, removeEmptyBlocksAroundCaret: removeEmptyBlocksAroundCaret,
    pastePlainText: pastePlainText, STYLE_MAP: STYLE_MAP, state: state
  };
  return API;
})(WW.util, WW.selection);
