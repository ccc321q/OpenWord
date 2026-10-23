/* ==========================================================================
   OpenWord — page layout: geometry, pagination, ruler, zoom, view modes
   Exposes WW.pages. The sheet is sized in CSS pixels at 96 dpi; the page
   model is {size, orientation, margins{top,right,bottom,left,header,footer}}
   in inches, shared with the DOCX/ODT exporters.
   ========================================================================== */
WW.pages = (function (U) {
  'use strict';

  var PXI = 96; // CSS pixels per inch

  var viewport, scaler, sheet, editor, header, footer, guides, rulerWrap, rulerCv;
  var boundaries = [];   // y offsets (layout px, relative to #editor content top) where a page starts
  var totalPages = 1;
  var lastCaretPage = 1;

  function page() { return WW.APP.page; }
  function sizeIn() {
    var s = WW.docxExport.PAGE_SIZES[page().size] || WW.docxExport.PAGE_SIZES.Letter;
    var w = s[0] / 1440, h = s[1] / 1440;
    if (page().orientation === 'landscape') { var t = w; w = h; h = t; }
    return [w, h];
  }

  function pageWidthPx() { return sizeIn()[0] * PXI; }
  function pageHeightPx() { return sizeIn()[1] * PXI; }
  function contentWidthPx() {
    var m = page().margins;
    return Math.max(120, (sizeIn()[0] - m.left - m.right) * PXI);
  }
  /** Usable height of one page's content column (between margins, minus header/footer). */
  function contentHeightPx() {
    var m = page().margins;
    var hfH = (header ? header.offsetHeight : 0) + (footer ? footer.offsetHeight : 0);
    return Math.max(200, pageHeightPx() - (m.top + m.bottom) * PXI - hfH);
  }

  /* ------------------------------------------------------------ apply -- */
  function applyPage() {
    if (!sheet) return;
    var m = page().margins;
    sheet.style.width = pageWidthPx() + 'px';
    sheet.style.minHeight = pageHeightPx() + 'px';
    sheet.style.padding = [m.top, m.right, m.bottom, m.left].map(function (v) {
      return (v * PXI) + 'px';
    }).join(' ');
    if (header) header.style.marginBottom = Math.max(0, (m.header - 0) * PXI - 8) + 'px';
    if (footer) footer.style.marginTop = Math.max(0, (m.footer - 0) * PXI - 8) + 'px';
    editor.style.minHeight = (contentHeightPx() - 8) + 'px';
    relayout();
  }

  /* ------------------------------------------------------- pagination -- */
  function computePages() {
    boundaries = [];
    if (!editor) { totalPages = 1; return; }
    var contentH = contentHeightPx();
    var kids = editor.children;
    var pageStart = 0;
    for (var i = 0; i < kids.length; i++) {
      var el = kids[i];
      var top = el.offsetTop, h = el.offsetHeight;
      if (top + h - pageStart > contentH) {
        if (h <= contentH) {
          boundaries.push(top);
          pageStart = top;
        } else {
          // block taller than a page: emit synthetic boundaries through it
          var n = Math.ceil((top + h - pageStart) / contentH);
          for (var k = 1; k < n; k++) boundaries.push(pageStart + k * contentH);
          pageStart = pageStart + (n - 1) * contentH;
        }
      }
      if (el.classList && el.classList.contains('ww-pagebreak')) {
        boundaries.push(top + h);
        pageStart = top + h;
      }
    }
    totalPages = Math.max(1, boundaries.length + 1);
  }

  function drawGuides() {
    if (!guides) return;
    guides.innerHTML = '';
    if (!document.body.classList.contains('show-guides')) return;
    var docTop = editor.offsetTop;
    boundaries.forEach(function (b, i) {
      var line = U.h('div', { class: 'page-break-line' }, [U.h('span', { text: 'Page ' + (i + 2) })]);
      line.style.top = (docTop + b) + 'px';
      guides.appendChild(line);
    });
  }

  /** Which page contains a layout-px y offset measured from the editor top. */
  function pageAtY(y) {
    var p = 1;
    for (var i = 0; i < boundaries.length; i++) {
      if (y >= boundaries[i]) p = i + 2; else break;
    }
    return U.clamp(p, 1, totalPages);
  }

  function caretYInEditor() {
    var sel = window.getSelection();
    if (!sel || !sel.rangeCount) return 0;
    var r = sel.getRangeAt(0);
    if (!editor.contains(r.startContainer)) return 0;
    var rect = r.getBoundingClientRect();
    var dRect = editor.getBoundingClientRect();
    var zoom = WW.APP.zoom || 1;
    return (rect.top - dRect.top) / zoom;
  }

  function pageAtCaret() {
    lastCaretPage = pageAtY(caretYInEditor());
    return lastCaretPage;
  }

  /* ----------------------------------------------------------- fields -- */
  function updateFields() {
    var now = new Date();
    U.$$('[data-ww-field]').forEach(function (el) {
      var kind = el.getAttribute('data-ww-field');
      var text = '';
      if (kind === 'page') text = String(lastCaretPage || 1);
      else if (kind === 'numpages') text = String(totalPages);
      else if (kind === 'date') text = now.toLocaleDateString();
      else if (kind === 'time') text = now.toLocaleTimeString();
      else if (kind === 'filename') text = WW.APP.docTitle || 'Untitled document';
      else if (kind === 'author') text = (WW.APP.meta && WW.APP.meta.author) || '';
      if (text && el.textContent !== text) el.textContent = text;
    });
  }

  /* ------------------------------------------------------------ ruler -- */
  function drawRuler() {
    if (!rulerCv || rulerWrap.hidden) return;
    var dpr = window.devicePixelRatio || 1;
    var w = rulerWrap.clientWidth, h = 24;
    if (rulerCv.width !== Math.round(w * dpr)) {
      rulerCv.width = Math.round(w * dpr);
      rulerCv.height = Math.round(h * dpr);
    }
    var ctx = rulerCv.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, w, h);

    var cs = getComputedStyle(document.body);
    var dim = cs.getPropertyValue('--text-dim') || '#605e5c';
    var line = cs.getPropertyValue('--border-strong') || '#c8c6c4';
    var bg = cs.getPropertyValue('--bg-chrome') || '#fff';

    var zoom = WW.APP.zoom || 1;
    var vRect = viewport.getBoundingClientRect();
    var sRect = sheet.getBoundingClientRect();
    var left = sRect.left - vRect.left;
    var m = page().margins;
    var cw = contentWidthPx() * zoom;
    var x0 = left + m.left * PXI * zoom;

    // page background + margin shading
    ctx.fillStyle = bg;
    ctx.fillRect(left, 0, sRect.width, h);
    ctx.fillStyle = 'rgba(128,128,128,.14)';
    ctx.fillRect(left, 0, m.left * PXI * zoom, h);
    ctx.fillRect(left + sRect.width - m.right * PXI * zoom, 0, m.right * PXI * zoom, h);

    // ticks
    ctx.strokeStyle = line;
    ctx.fillStyle = dim;
    ctx.font = '9px ' + (cs.getPropertyValue('--font-ui') || 'sans-serif');
    ctx.textBaseline = 'top';
    var inches = Math.floor(contentWidthPx() / PXI);
    for (var i = 0; i <= inches; i++) {
      var x = x0 + i * PXI * zoom;
      ctx.beginPath();
      ctx.moveTo(x + .5, h - 9);
      ctx.lineTo(x + .5, h);
      ctx.stroke();
      ctx.fillText(String(i), x + 2, 1);
      // half-inch ticks
      var xh = x0 + (i + .5) * PXI * zoom;
      if (i < inches) {
        ctx.beginPath();
        ctx.moveTo(xh + .5, h - 5);
        ctx.lineTo(xh + .5, h);
        ctx.stroke();
      }
    }
    ctx.strokeStyle = line;
    ctx.strokeRect(left + .5, .5, sRect.width - 1, h - 1);

    // indent markers for the caret block
    var b = WW.commands && WW.commands.anchorBlock();
    if (b && editor.contains(b)) {
      var bcs = window.getComputedStyle(b);
      var ml = (parseFloat(bcs.marginLeft) || 0) + (parseFloat(bcs.textIndent) || 0);
      var mr = parseFloat(bcs.marginRight) || 0;
      drawMarker(ctx, x0 + ml * zoom, h - 1, '#2b579a', true);
      drawMarker(ctx, x0 + cw - mr * zoom, h - 1, '#2b579a', false);
    }
  }

  function drawMarker(ctx, x, y, color, leftSide) {
    ctx.fillStyle = color;
    ctx.beginPath();
    if (leftSide) { ctx.moveTo(x, y - 8); ctx.lineTo(x + 6, y - 8); ctx.lineTo(x, y); }
    else { ctx.moveTo(x, y - 8); ctx.lineTo(x - 6, y - 8); ctx.lineTo(x, y); }
    ctx.closePath();
    ctx.fill();
  }

  /* ------------------------------------------------------------- zoom -- */
  function setZoom(z, fromSlider) {
    z = U.clamp(z, 0.5, 2.5);
    WW.APP.zoom = z;
    scaler.style.zoom = String(z);
    var pct = Math.round(z * 100) + '%';
    U.$('#zoom-label').textContent = pct;
    U.$('#status-zoom').textContent = pct;
    var range = U.$('#zoom-range');
    if (!fromSlider) range.value = String(Math.round(z * 100));
    drawRuler();
    if (WW.storage) WW.storage.setPref('zoom', z);
  }

  /* -------------------------------------------------------- view mode -- */
  function setViewMode(mode) {
    document.body.classList.remove('view-web', 'view-focus');
    if (mode === 'web') document.body.classList.add('view-web');
    if (mode === 'focus') document.body.classList.add('view-focus');
    U.$('#view-mode').value = mode;
    U.$('#status-view').value = mode;
    WW.APP.viewMode = mode;
    markFocusBlock();
    relayout();
    if (WW.storage) WW.storage.setPref('view', mode);
  }

  function markFocusBlock() {
    U.$$('.doc > .ww-focus', editor).forEach(function (el) { el.classList.remove('ww-focus'); });
    if (!document.body.classList.contains('view-focus')) return;
    var b = WW.commands && WW.commands.anchorBlock();
    while (b && b.parentElement !== editor) b = b.parentElement;
    if (b) b.classList.add('ww-focus');
  }

  /* ---------------------------------------------------------- toggles -- */
  function setRuler(on) {
    rulerWrap.hidden = !on;
    if (on) drawRuler();
    if (WW.storage) WW.storage.setPref('ruler', !!on);
  }
  function setMarks(on) {
    document.body.classList.toggle('show-marks', !!on);
    if (WW.storage) WW.storage.setPref('marks', !!on);
  }
  function setGuides(on) {
    document.body.classList.toggle('show-guides', !!on);
    drawGuides();
    if (WW.storage) WW.storage.setPref('guides', !!on);
  }

  /* --------------------------------------------------------- relayout -- */
  var pending = false;
  function relayout() {
    if (pending) return;
    pending = true;
    requestAnimationFrame(function () {
      pending = false;
      computePages();
      drawGuides();
      updateFields();
      drawRuler();
      if (WW.app) WW.app.updateStatus();
    });
  }

  /* ------------------------------------------------------------- init -- */
  function init() {
    viewport = U.$('#viewport');
    scaler = U.$('#sheet-scaler');
    sheet = U.$('#sheet');
    editor = U.$('#editor');
    header = U.$('#page-header');
    footer = U.$('#page-footer');
    guides = U.$('#page-guides');
    rulerWrap = U.$('#ruler-wrap');
    rulerCv = U.$('#ruler');

    applyPage();
    window.addEventListener('resize', U.debounce(function () { drawRuler(); drawGuides(); }, 120));
    viewport.addEventListener('scroll', U.throttle(drawRuler, 100));
  }

  return {
    init: init, applyPage: applyPage, relayout: relayout,
    contentWidthPx: contentWidthPx, contentHeightPx: contentHeightPx,
    pageWidthPx: pageWidthPx, pageHeightPx: pageHeightPx,
    totalPages: function () { return totalPages; },
    pageAtCaret: pageAtCaret, updateFields: updateFields,
    setZoom: setZoom, setViewMode: setViewMode, markFocusBlock: markFocusBlock,
    setRuler: setRuler, setMarks: setMarks, setGuides: setGuides,
    drawRuler: drawRuler, PXI: PXI
  };
})(WW.util);
