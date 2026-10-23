/* ==========================================================================
   OpenWord — HTML import / export, plain text, and document normalisation
   Handles the messy HTML produced by Word ("Save as Word HTML") and by
   Google Docs ("Download as HTML"), both of which style everything through
   generated CSS classes.
   ========================================================================== */
WW.htmlio = (function (U) {
  'use strict';

  var KEEP_TAGS = [
    'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li', 'b', 'strong', 'i', 'em', 'u', 'ins',
    's', 'strike', 'del', 'sub', 'sup', 'span', 'a', 'img', 'br', 'hr', 'table', 'thead', 'tbody',
    'tfoot', 'tr', 'td', 'th', 'colgroup', 'col', 'blockquote', 'pre', 'code', 'kbd', 'samp', 'small',
    'mark', 'abbr', 'q', 'cite', 'figure', 'figcaption', 'div', 'font', 'big', 'label', 'time', 'data',
    'ruby', 'rt', 'rp', 'bdi', 'bdo', 'article', 'section', 'header', 'footer', 'main', 'nav'
  ];

  var KEEP_ATTRS = [
    'style', 'href', 'title', 'src', 'alt', 'width', 'height', 'colspan', 'rowspan', 'class',
    'data-ww-list-type', 'data-ww-field', 'data-ww-value', 'data-ww-seq', 'data-ww-key',
    'data-ww-marker', 'data-borders', 'data-header', 'data-ww-synthetic', 'data-ww-indent'
  ];

  /* ------------------------------------------------------ sanitiser ----- */
  function sanitize(html) {
    var cfg = {
      ALLOWED_TAGS: KEEP_TAGS,
      ALLOWED_ATTR: KEEP_ATTRS,
      ALLOW_DATA_ATTR: true,
      KEEP_CONTENT: true,
      RETURN_DOM: false,
      FORBID_TAGS: ['script', 'style', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'select', 'textarea', 'link', 'meta', 'base', 'svg', 'math', 'noscript'],
      FORBID_ATTR: ['onerror', 'onload', 'onclick', 'onmouseover', 'srcset', 'background', 'dynsrc', 'lowsrc', 'contenteditable', 'tabindex', 'id', 'name', 'lang', 'dir', 'align', 'valign', 'bgcolor', 'border', 'cellpadding', 'cellspacing', 'face', 'size', 'color']
    };
    if (window.DOMPurify && DOMPurify.sanitize) {
      DOMPurify.setConfig(cfg);
      var out = DOMPurify.sanitize(html);
      DOMPurify.clearConfig();
      return out;
    }
    // Very small fallback if the vendored library failed to load.
    var div = document.createElement('div');
    div.innerHTML = html;
    U.$$('script,style,iframe,object,embed,form,link,meta,svg', div).forEach(function (n) { n.remove(); });
    return div.innerHTML;
  }

  /* --------------------------------------------------- CSS class rules -- */
  /**
   * Parse the <style> blocks of an exported document and inline the declarations
   * onto the elements that match, because Word/Docs keep all formatting there.
   */
  function inlineCssRules(doc) {
    var rules = [];
    U.$$('style', doc).forEach(function (styleEl) {
      var text = styleEl.textContent || '';
      text = text.replace(/\/\*[\s\S]*?\*\//g, '');
      var ruleRe = /([^{}]+)\{([^{}]*)\}/g;
      var m;
      while ((m = ruleRe.exec(text))) {
        var selectors = m[1].split(',').map(function (s) { return s.trim(); }).filter(Boolean);
        var decl = parseDeclarations(m[2]);
        if (!Object.keys(decl).length) continue;
        selectors.forEach(function (sel) {
          if (/@/.test(sel)) return;
          rules.push({ selector: sel, decl: decl });
        });
      }
    });
    if (!rules.length) return rules;

    // Apply rules from least to most specific so later ones win.
    rules.sort(function (a, b) { return specificity(a.selector) - specificity(b.selector); });
    rules.forEach(function (rule) {
      var els;
      try { els = doc.querySelectorAll(rule.selector); } catch (e) { els = []; }
      Array.prototype.forEach.call(els, function (el) {
        Object.keys(rule.decl).forEach(function (prop) {
          if (!INTERESTING[prop]) return;
          var value = rule.decl[prop];
          if (/^mso-|inherit|auto$/i.test(prop) || value === 'none' && prop === 'text-decoration') { /* fall through, still set text-decoration */ }
          if (prop.indexOf('mso-') === 0) return;
          try { el.style.setProperty(prop, value); } catch (e) { /* ignore */ }
        });
      });
    });
    U.$$('style', doc).forEach(function (s) { s.remove(); });
    return rules;
  }

  var INTERESTING = {
    'font-weight': 1, 'font-style': 1, 'font-size': 1, 'font-family': 1, 'color': 1,
    'background-color': 1, 'text-decoration': 1, 'text-decoration-line': 1, 'text-align': 1,
    'margin': 1, 'margin-top': 1, 'margin-bottom': 1, 'margin-left': 1, 'margin-right': 1,
    'text-indent': 1, 'line-height': 1, 'list-style-type': 1, 'padding-left': 1, 'vertical-align': 1,
    'border-collapse': 1, 'border': 1, 'border-top': 1, 'border-bottom': 1, 'border-left': 1, 'border-right': 1,
    'border-color': 1, 'border-style': 1, 'width': 1, 'height': 1, 'white-space': 1
  };

  function specificity(sel) {
    var n = 0;
    (sel.match(/\./g) || []).forEach(function () { n += 10; });
    (sel.match(/#/g) || []).forEach(function () { n += 100; });
    (sel.match(/[a-z]+$/i) || []).forEach(function () { n += 1; });
    return n;
  }

  function parseDeclarations(text) {
    var out = {};
    String(text).split(';').forEach(function (chunk) {
      var i = chunk.indexOf(':');
      if (i < 0) return;
      var prop = chunk.slice(0, i).trim().toLowerCase();
      var value = chunk.slice(i + 1).trim().replace(/\s*!important$/, '');
      if (!prop || !value) return;
      out[prop] = value;
    });
    return out;
  }

  /* ------------------------------------------------- Google Docs tweaks -- */
  function looksLikeGoogleHtml(doc, raw) {
    return /docs-dlp|kix|lst-kix|docs-controller|id="docs-internal-guid/i.test(raw || '') ||
      !!doc.querySelector('[class*="lst-kix"], .docs-dlp, #docs-internal-guid');
  }

  function tweakGoogleHtml(doc) {
    // Google nests the real content inside a guid container.
    var guid = doc.querySelector('#docs-internal-guid');
    var root = guid || doc.body || doc.documentElement;

    // ul/li > p : lift paragraph content into the list item
    U.$$('li p', root).forEach(function (p) {
      var li = p.parentElement;
      if (!li || li.tagName !== 'LI') return;
      if (p.style.textAlign) li.style.textAlign = li.style.textAlign || p.style.textAlign;
      var frag = document.createDocumentFragment();
      Array.prototype.forEach.call(p.childNodes, function (n) { frag.appendChild(n.cloneNode(true)); });
      p.replaceWith(frag);
    });

    // list-type from CSS decides <ul> vs <ol>
    U.$$('ul', root).forEach(function (ul) {
      var st = (ul.style.listStyleType || '').toLowerCase();
      var ordered = /^(decimal|lower-alpha|upper-alpha|lower-roman|upper-roman|decimal-leading-zero|lower-latin|upper-latin)/.test(st);
      if (ordered) {
        var ol = document.createElement('ol');
        copyAttrs(ul, ol);
        ol.style.listStyleType = st;
        ol.dataset.wwListType = cssTypeToInternal(st);
        ol.innerHTML = ul.innerHTML;
        ul.replaceWith(ol);
      } else if (st && !/^(disc|circle|square|none)/.test(st)) {
        ul.style.listStyleType = '"'+st.replace(/["']/g, '') + '"';
      }
    });

    // Google writes headings as big bold paragraphs when exported without tags
    U.$$('p', root).forEach(function (p) {
      if (p.closest('li, table')) return;
      var text = (p.textContent || '').trim();
      if (!text || text.length > 160) return;
      var spans = p.querySelectorAll('span, b');
      var big = 0, bold = 0, total = 0;
      Array.prototype.forEach.call(spans.length ? spans : [p], function (el) {
        var cs = el.style;
        var fs = parseFloat(cs.fontSize);
        total++;
        if (fs >= 13) big++;
        if (cs.fontWeight === 'bold' || parseInt(cs.fontWeight, 10) >= 600) bold++;
      });
      if (!total) return;
      if (bold / total >= 0.6 && big / total >= 0.6) {
        var maxSize = 0;
        Array.prototype.forEach.call(spans.length ? spans : [p], function (el) {
          var fs = parseFloat(el.style.fontSize) || 0;
          if (fs > maxSize) maxSize = fs;
        });
        var level = maxSize >= 18 ? 1 : maxSize >= 14.5 ? 2 : 3;
        var h = document.createElement('h' + level);
        copyAttrs(p, h);
        h.innerHTML = p.innerHTML.replace(/<(b|strong)[^>]*>([\s\S]*?)<\/\1>/gi, '$2');
        p.replaceWith(h);
      }
    });
    return root;
  }

  function cssTypeToInternal(st) {
    st = String(st).toLowerCase();
    if (/lower-alpha|lower-latin/.test(st)) return 'lower-alpha';
    if (/upper-alpha|upper-latin/.test(st)) return 'upper-alpha';
    if (/lower-roman/.test(st)) return 'lower-roman';
    if (/upper-roman/.test(st)) return 'upper-roman';
    if (/decimal-leading-zero/.test(st)) return 'decimal-zero';
    return '';
  }

  function copyAttrs(from, to) {
    Array.prototype.forEach.call(from.attributes, function (a) {
      if (a.name === 'class' && /^(lst-|c\d)/.test(a.value)) return;
      if (a.name.indexOf('data-') === 0 || KEEP_ATTRS.indexOf(a.name) >= 0) {
        try { to.setAttribute(a.name, a.value); } catch (e) { }
      }
    });
  }

  /* ------------------------------------------------- Word HTML tweaks --- */
  function tweakWordHtml(doc) {
    // Word's Word-processing ML leftovers
    U.$$('o\\:p, o\\:span, w\\:any, [class*="MsoIns"], [style*="mso-"]', doc).forEach(function (el) {
      if (el.tagName.toLowerCase() === 'o:p' && (el.getAttribute('class') || '').match(/List\d+/i)) {
        // paragraph marker inside a list item; drop it
        el.remove();
        return;
      }
    });
    U.$$('[style]', doc).forEach(function (el) {
      var style = el.getAttribute('style') || '';
      var cleaned = style.replace(/mso-[^;]+;?/gi, '').replace(/tab-stops:[^;]+;?/gi, '');
      if (cleaned.trim()) el.setAttribute('style', cleaned);
      else el.removeAttribute('style');
    });
    U.$$('[class]', doc).forEach(function (el) {
      var cls = (el.getAttribute('class') || '').replace(/\bMso[A-Za-z]+\b/g, '').replace(/\bSpellE\b/g, '').trim();
      if (cls) el.setAttribute('class', cls); else el.removeAttribute('class');
    });
  }

  /* ------------------------------------------------------ normalise ----- */
  var FONT_SIZES = { 1: '8pt', 2: '10pt', 3: '12pt', 4: '14pt', 5: '18pt', 6: '24pt', 7: '36pt' };
  var ABSOLUTE_SIZES = {
    'xx-small': '7pt', 'x-small': '8pt', 'small': '10pt', 'medium': '12pt',
    'large': '14pt', 'x-large': '18pt', 'xx-large': '24pt', 'xxx-large': '32pt', 'larger': '12pt', 'smaller': '10pt'
  };

  function normalizeFragment(root, opts) {
    opts = opts || {};
    // 1. <font> → span / b / i / u
    U.$$('font', root).forEach(function (f) {
      var style = '';
      var face = f.getAttribute('face'), size = f.getAttribute('size'), color = f.getAttribute('color');
      if (face) style += 'font-family:' + face + ';';
      if (size) style += 'font-size:' + (FONT_SIZES[size] || size) + ';';
      if (color) style += 'color:' + (/^[0-9a-f]{6}$/i.test(color) ? '#' + color : color) + ';';
      if (style) f.setAttribute('style', style);
      f.removeAttribute('face'); f.removeAttribute('size'); f.removeAttribute('color');
      unwrap(f);
    });

    // 2. semantic inline tags → canonical set
    renameTags(root, { strong: 'b', em: 'i', strike: 's', ins: 'u' });

    // 3. absolute font sizes → pt
    U.$$('[style*="font-size"]', root).forEach(function (el) {
      var v = el.style.fontSize;
      if (!v) return;
      var key = v.trim().toLowerCase();
      if (ABSOLUTE_SIZES[key]) el.style.fontSize = ABSOLUTE_SIZES[key];
      else if (/^[\d.]+$/.test(key)) el.style.fontSize = Math.round(parseFloat(key) * 0.75 * 100) / 100 + 'pt';
    });

    // 4. text-decoration shorthands coming from execCommand
    U.$$('[style]', root).forEach(function (el) {
      var td = (el.style.textDecorationLine || el.style.textDecoration || '').toLowerCase();
      if (!td) return;
      var style = el.getAttribute('style') || '';
      style = style.replace(/text-decoration(-line)?:[^;]+;?/gi, '');
      var deco = [];
      if (td.indexOf('underline') >= 0) deco.push('underline');
      if (td.indexOf('line-through') >= 0) deco.push('line-through');
      if (deco.length) style += 'text-decoration:' + deco.join(' ') + ';';
      if (style.trim()) el.setAttribute('style', style.trim()); else el.removeAttribute('style');
    });

    // 5. <b>/<i>/<u>/<s> carrying only styling → keep tags, drop redundant styles
    U.$$('b[style*="font-weight: normal"], i[style*="font-style: normal"], u[style*="text-decoration: none"], s[style*="text-decoration: none"]', root)
      .forEach(function (el) { unwrap(el); });

    // 6. spans with no style and no data are noise
    U.$$('span', root).forEach(function (s) {
      if (!s.getAttribute('style') && !Array.prototype.some.call(s.attributes, function (a) { return a.name.indexOf('data-') === 0; })) unwrap(s);
    });

    // 7. block-level clean-up: divs become paragraphs unless they wrap blocks
    var divs = U.$$('div', root);
    divs.forEach(function (d) {
      var hasBlocks = Array.prototype.some.call(d.children, function (c) { return isBlockTag(c.tagName); });
      if (hasBlocks) { unwrap(d); return; }
      var tag = d.tagName.toLowerCase();
      if (tag !== 'div') return;
      var p = document.createElement('p');
      copyAttrs(d, p);
      var cls = (p.getAttribute('class') || '').replace(/\b(WordSection\d+|ExternalClass\S*)\b/g, '').trim();
      if (cls) p.setAttribute('class', cls); else p.removeAttribute('class');
      p.innerHTML = d.innerHTML;
      d.replaceWith(p);
    });

    // 8. blocks directly inside <li> collapse when they are the only child
    U.$$('li', root).forEach(function (li) {
      if (li.children.length === 1 && /^(P|DIV)$/.test(li.children[0].tagName)) {
        var p = li.children[0];
        var frag = document.createDocumentFragment();
        Array.prototype.forEach.call(p.childNodes, function (n) { frag.appendChild(n.cloneNode(true)); });
        li.innerHTML = '';
        li.appendChild(frag);
      }
      ensureFirstBlockIsParagraph(li);
    });

    // 9. tables: blockify cell content, drop nested layout junk
    U.$$('table', root).forEach(function (t) {
      Array.prototype.slice.call(t.rows).forEach(function (row) {
        if (!row.cells.length) { row.remove(); return; }
        Array.prototype.slice.call(row.cells).forEach(function (cell) {
          if (cell.tagName === 'TH') { /* keep */ }
          var inline = !Array.prototype.some.call(cell.childNodes, function (n) {
            return n.nodeType === 1 && isBlockTag(n.tagName);
          });
          if (inline) {
            var p = document.createElement('p');
            while (cell.firstChild) p.appendChild(cell.firstChild);
            if (!cell.textContent.trim() && !cell.querySelector('img')) p.appendChild(document.createTextNode(''));
            cell.appendChild(p);
          }
        });
      });
      if (!t.style.width) {
        var measured = t.getBoundingClientRect ? 0 : 0;
        if (measured) t.style.width = measured + 'px';
      }
    });

    // 10. links: require href, strip anchors
    U.$$('a', root).forEach(function (a) {
      var href = a.getAttribute('href');
      if (!href || href.charAt(0) === '#' || /^(javascript|data):/i.test(href)) { unwrap(a); return; }
      a.removeAttribute('target');
    });

    // 11. images: keep them, ensure sane sizing
    U.$$('img', root).forEach(function (img) {
      var src = img.getAttribute('src') || '';
      if (!src || /^(javascript|about):/i.test(src)) { img.remove(); return; }
      if (img.style.width && parseFloat(img.style.width) > 700) {
        var ratio = (parseFloat(img.style.height) || 0) / (parseFloat(img.style.width) || 1);
        img.style.height = Math.round(700 * ratio) + 'px';
        img.style.width = '700px';
      }
    });

    // 12. remove nested/empty leftovers
    U.$$('p > p, div > div[class=""]', root).forEach(function (p) { unwrap(p); });
    U.$$('li:empty, p:empty', root).forEach(function () { /* empty paragraphs are meaningful */ });

    // 13. headings must not contain block children
    U.$$('h1,h2,h3,h4,h5,h6', root).forEach(function (h) {
      if (Array.prototype.some.call(h.children, function (c) { return isBlockTag(c.tagName); })) {
        var p = document.createElement('p');
        copyAttrs(h, p);
        p.innerHTML = h.innerHTML.replace(/<\/?(p|div)[^>]*>/gi, '');
        h.replaceWith(p);
      }
      if (!h.textContent.trim()) h.appendChild(document.createElement('br'));
    });

    // 14. top and tail: the editor root always starts and ends with a block
    while (root.firstChild && root.firstChild.nodeType === 3 && !root.firstChild.nodeValue.trim()) root.removeChild(root.firstChild);
    if (root.firstChild && root.firstChild.nodeType === 3) {
      var p0 = document.createElement('p');
      while (root.firstChild && root.firstChild.nodeType === 3) p0.appendChild(root.firstChild);
      root.insertBefore(p0, root.firstChild);
    }
    if (!root.firstChild) root.appendChild(document.createElement('p'));
    var last = root.lastElementChild;
    if (!last || !isBlockTag(last.tagName)) {
      var pEnd = document.createElement('p');
      root.appendChild(pEnd);
    }
    // Drop stray whitespace text nodes between blocks (they create phantom lines)
    Array.prototype.slice.call(root.childNodes).forEach(function (n) {
      if (n.nodeType === 3 && !n.nodeValue.trim()) root.removeChild(n);
    });
    return root;
  }

  function isBlockTag(tagName) {
    return ['P', 'DIV', 'UL', 'OL', 'TABLE', 'BLOCKQUOTE', 'PRE', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'HR', 'FIGURE'].indexOf(tagName) >= 0;
  }

  function ensureFirstBlockIsParagraph(li) {
    if (!li.firstChild) { li.appendChild(document.createElement('p')); return; }
    if (li.firstChild.nodeType === 3 || !isBlockTag(li.firstChild.tagName)) {
      var p = document.createElement('p');
      var frag = document.createDocumentFragment();
      Array.prototype.slice.call(li.childNodes).forEach(function (n) {
        if (n.nodeType === 1 && isBlockTag(n.tagName)) return;
        frag.appendChild(n);
      });
      p.appendChild(frag);
      li.insertBefore(p, li.firstChild);
    }
  }

  function unwrap(el) {
    var parent = el.parentNode;
    if (!parent) return;
    while (el.firstChild) parent.insertBefore(el.firstChild, el);
    parent.removeChild(el);
  }

  function renameTags(root, map) {
    Object.keys(map).forEach(function (from) {
      U.$$(from, root).forEach(function (el) {
        var rep = document.createElement(map[from]);
        Array.prototype.forEach.call(el.attributes, function (a) {
          try { rep.setAttribute(a.name, a.value); } catch (e) { }
        });
        rep.innerHTML = el.innerHTML;
        el.replaceWith(rep);
      });
    });
  }

  /* ------------------------------------------------------ entry points -- */
  /** HTML file / fragment in, editor-ready HTML out. */
  function fragmentFromHtml(html, options) {
    options = options || {};
    var clean = sanitize(html);
    var host = document.createElement('div');
    host.innerHTML = clean;

    var fullDoc = /<html[\s>]|<!DOCTYPE/i.test(html);
    var parsed = null;
    if (fullDoc) {
      parsed = new DOMParser().parseFromString(html, 'text/html');
      inlineCssRules(parsed);
      var isGoogle = looksLikeGoogleHtml(parsed, html);
      var root = tweakGoogleHtml(parsed);
      var container = document.createElement('div');
      Array.prototype.slice.call(root.childNodes).forEach(function (n) {
        if (n.nodeType === 1 && /^(SCRIPT|STYLE|LINK|META|TITLE|HEAD)$/.test(n.tagName)) return;
        container.appendChild(n.cloneNode(true));
      });
      if (!container.childNodes.length && parsed.body) container.innerHTML = parsed.body.innerHTML;
      tweakWordHtml(container);
      normalizeFragment(container, options);
      container._source = isGoogle ? 'google' : 'html';
      return container;
    }

    // Fragment: still inline any <style> the user pasted along with markup
    var styleHost = document.createElement('div');
    styleHost.innerHTML = html;
    inlineCssRules(styleHost);
    tweakWordHtml(host);
    normalizeFragment(host, options);
    host._source = 'html';
    return host;
  }

  function fromPlainText(text) {
    var host = document.createElement('div');
    var lines = String(text).replace(/\r\n?/g, '\n').split('\n\n+');
    if (!lines.length) host.appendChild(document.createElement('p'));
    lines.forEach(function (para) {
      var p = document.createElement('p');
      var parts = para.split('\n');
      parts.forEach(function (line, i) {
        if (i) p.appendChild(document.createElement('br'));
        p.appendChild(document.createTextNode(line));
      });
      host.appendChild(p);
    });
    normalizeFragment(host);
    return host;
  }

  /** Editor DOM -> a complete, self-contained HTML document. */
  function toHtmlDocument(editor, opts) {
    opts = opts || {};
    var title = opts.title || 'Document';
    var body = editor.innerHTML
      .replace(/ class="ww-focus"/g, '')
      .replace(/<span[^>]*class="ww-caret"[^>]*>[\u200b]*<\/span>/g, '');
    var hfCss = '';
    var header = opts.header && U.stripZwsp(opts.header.textContent || '').trim() ?
      '<div class="doc-header">' + opts.header.innerHTML + '</div>' : '';
    var footer = opts.footer && U.stripZwsp(opts.footer.textContent || '').trim() ?
      '<div class="doc-footer">' + opts.footer.innerHTML + '</div>' : '';
    var page = opts.page || { size: 'Letter', orientation: 'portrait', margins: { top: 1, right: 1, bottom: 1, left: 1 } };
    var size = WW.docxExport.PAGE_SIZES[page.size] || WW.docxExport.PAGE_SIZES.Letter;
    var wIn = size[0] / 1440, hIn = size[1] / 1440;
    if (page.orientation === 'landscape') { var t = wIn; wIn = hIn; hIn = t; }
    return '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<title>' + U.escapeHtml(title) + '</title>\n' +
      '<style>\n' +
      '@page { size: ' + wIn.toFixed(2) + 'in ' + hIn.toFixed(2) + 'in; margin: ' +
      [page.margins.top, page.margins.right, page.margins.bottom, page.margins.left].map(function (m) { return m + 'in'; }).join(' ') + '; }\n' +
      'body { font-family: Calibri, Carlito, "Segoe UI", Arial, sans-serif; font-size: 11pt; line-height: 1.5; color: #111; max-width: ' + (wIn - page.margins.left - page.margins.right).toFixed(2) + 'in; margin: 1in auto; }\n' +
      'p, h1, h2, h3, h4, h5, h6, li { white-space: pre-wrap; overflow-wrap: break-word; }\n' +
      'p { margin: 0 0 8pt; } h1 { font-size: 20pt; color: #2b579a; } h2 { font-size: 16pt; color: #2b579a; } h3 { font-size: 13.5pt; color: #3b5b8c; }\n' +
      '.ww-title { font-size: 26pt; color: #2b579a; } .ww-subtitle { font-size: 13pt; font-style: italic; color: #5b7a9d; }\n' +
      '.ww-quote { font-style: italic; margin-left: 24pt; } blockquote { font-style: italic; margin: 8pt 24pt; }\n' +
      'table { border-collapse: collapse; } td, th { border: 1px solid #ababab; padding: 4pt 6pt; vertical-align: top; }\n' +
      'a { color: #0563c1; } img { max-width: 100%; }\n' +
      '.ww-pagebreak { page-break-after: always; break-after: page; height: 0; }\n' +
      '.ww-field { background: #e8eef8; padding: 0 2px; }\n' +
      '.doc-header, .doc-footer { color: #666; font-size: 9pt; border-bottom: 1px solid #ddd; }\n' +
      '.doc-footer { border-bottom: 0; border-top: 1px solid #ddd; }\n' +
      hfCss +
      '</style>\n</head>\n<body>\n' +
      (header || footer ? '<header>' + header + '</header>\n' : '') +
      body + '\n' +
      (footer ? '<footer>' + footer + '</footer>\n' : '') +
      '<!-- Created with OpenWord -->\n</body>\n</html>';
  }

  function toPlainText(el) {
    var out = [];
    (function walk(node, depth) {
      for (var i = 0; i < node.childNodes.length; i++) {
        var c = node.childNodes[i];
        if (c.nodeType === 3) { out.push(c.nodeValue); continue; }
        if (c.nodeType !== 1) continue;
        var tag = c.tagName;
        if (tag === 'BR') { out.push('\n'); continue; }
        if (tag === 'HR') { out.push('\n----------------------------------------\n'); continue; }
        if (tag === 'IMG') { out.push(c.getAttribute('alt') ? '[' + c.getAttribute('alt') + ']' : ''); continue; }
        if (tag === 'LI') {
          out.push('\n' + new Array(depth + 1).join('  ') + '\u2022 ');
          walk(c, depth + 1);
          continue;
        }
        if (isBlockTag(tag)) {
          out.push('\n');
          walk(c, depth);
          out.push('\n');
          continue;
        }
        walk(c, depth);
      }
    })(el, 0);
    return out.join('')
      .replace(/[\u00a0]/g, ' ')
      .replace(/\u200b|\ufeff/g, '')
      .replace(/\n{3,}/g, '\n\n')
      .trim() + '\n';
  }

  return {
    sanitize: sanitize, fragmentFromHtml: fragmentFromHtml, fromPlainText: fromPlainText,
    toHtmlDocument: toHtmlDocument, toPlainText: toPlainText, normalizeFragment: normalizeFragment,
    inlineCssRules: inlineCssRules, looksLikeGoogleHtml: looksLikeGoogleHtml
  };
})(WW.util);
