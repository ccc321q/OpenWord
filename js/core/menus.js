/* ==========================================================================
   OpenWord — menu bar dropdowns and the right-click context menu.
   Exposes WW.menus. Menu items are declarative:
     { label, sc, cmd, value, dialog, action, check(), disabled(), sep, head }
   Commands run through WW.app.runCmd so app-level and editor commands share
   one dispatcher; checks are refreshed every time a menu opens.
   ========================================================================== */
WW.menus = (function (U) {
  'use strict';

  var menubar, ctxMenu, openMenu = null;

  function run(item) {
    closeAll();
    if (item.action) { item.action(); return; }
    if (item.dialog) { WW.dialogs.open(item.dialog); return; }
    if (item.cmd) WW.app.runCmd(item.cmd, item.value);
  }

  function canUndo() { return !!(WW.APP && WW.APP.canUndo); }
  function canRedo() { return !!(WW.APP && WW.APP.canRedo); }
  function qs(cmd) { return WW.commands.queryState(cmd); }

  /* ------------------------------------------------------------- spec -- */
  function styleItems() {
    return [
      ['Normal', 'Normal'], ['Heading 1', 'Heading1'], ['Heading 2', 'Heading2'],
      ['Heading 3', 'Heading3'], ['Title', 'Title'], ['Subtitle', 'Subtitle'],
      ['Quote', 'Quote'], ['Intense Quote', 'IntenseQuote'], ['Code', 'Code']
    ].map(function (s) {
      return {
        label: s[0], cmd: 'formatBlock', value: s[1],
        check: function () { return WW.commands.currentBlockStyle() === s[1]; }
      };
    });
  }

  function recentItems() {
    var list = WW.storage.getRecent();
    if (!list.length) return [{ head: 'Recent' }, { empty: true }];
    return [{ head: 'Recent' }].concat(list.map(function (r) {
      return { label: r.name, action: function () { WW.app.openSampleByName(r.name); } };
    }));
  }

  function spec() {
    return [
      { id: 'file', label: 'File', items: [
        { label: 'New', sc: 'Ctrl+Alt+N', cmd: 'new' },
        { label: 'Open…', sc: 'Ctrl+O', cmd: 'open' },
        { sep: true }, { head: 'Save & export' },
        { label: 'Save as .docx', sc: 'Ctrl+S', cmd: 'save-docx' },
        { label: 'Save for Google Docs…', cmd: 'save-gdoc' },
        { label: 'Save as .html', cmd: 'save-html' },
        { label: 'Save as .txt', cmd: 'save-txt' },
        { label: 'Save as .odt', cmd: 'save-odt' },
        { label: 'Print / PDF', sc: 'Ctrl+P', cmd: 'print' },
        { sep: true }
      ].concat(recentItems()).concat([
        { sep: true },
        { label: 'Clear autosave', cmd: 'clear-autosave' }
      ]) },

      { id: 'edit', label: 'Edit', items: [
        { label: 'Undo', sc: 'Ctrl+Z', action: function () { WW.history.undo(); }, disabled: function () { return !canUndo(); } },
        { label: 'Redo', sc: 'Ctrl+Y', action: function () { WW.history.redo(); }, disabled: function () { return !canRedo(); } },
        { sep: true },
        { label: 'Cut', sc: 'Ctrl+X', cmd: 'cut' },
        { label: 'Copy', sc: 'Ctrl+C', cmd: 'copy' },
        { label: 'Paste', sc: 'Ctrl+V', cmd: 'paste' },
        { label: 'Paste plain text', sc: 'Ctrl+Shift+V', cmd: 'paste-plain' },
        { sep: true },
        { label: 'Find & Replace…', sc: 'Ctrl+F', action: function () { WW.app.openFind(false); } },
        { label: 'Replace…', sc: 'Ctrl+H', action: function () { WW.app.openFind(true); } },
        { label: 'Select all', sc: 'Ctrl+A', cmd: 'select-all' },
        { sep: true },
        { label: 'Format painter', sc: 'Ctrl+Alt+C', cmd: 'format-painter' }
      ] },

      { id: 'view', label: 'View', items: [
        { label: 'Ruler', check: function () { return !U.$('#ruler-wrap').hidden; }, action: function () { WW.app.runCmd('toggle-ruler'); } },
        { label: 'Formatting marks', check: function () { return document.body.classList.contains('show-marks'); }, action: function () { WW.app.runCmd('toggle-marks'); } },
        { label: 'Page guides', check: function () { return document.body.classList.contains('show-guides'); }, action: function () { WW.app.runCmd('toggle-grid'); } },
        { sep: true }, { head: 'View mode' },
        { label: 'Print layout', check: function () { return WW.APP.viewMode === 'print'; }, action: function () { WW.pages.setViewMode('print'); } },
        { label: 'Web layout', check: function () { return WW.APP.viewMode === 'web'; }, action: function () { WW.pages.setViewMode('web'); } },
        { label: 'Focus', check: function () { return WW.APP.viewMode === 'focus'; }, action: function () { WW.pages.setViewMode('focus'); } },
        { sep: true },
        { label: 'Zoom in', action: function () { WW.app.runCmd('zoom-in'); } },
        { label: 'Zoom out', action: function () { WW.app.runCmd('zoom-out'); } },
        { label: 'Zoom reset', action: function () { WW.app.runCmd('zoom-reset'); } },
        { sep: true },
        { label: 'Dark theme', check: function () { return document.body.classList.contains('theme-dark'); }, action: function () { WW.app.runCmd('toggle-theme'); } }
      ] },

      { id: 'insert', label: 'Insert', items: [
        { label: 'Table…', dialog: 'table' },
        { label: 'Picture…', dialog: 'image' },
        { label: 'Link…', sc: 'Ctrl+K', dialog: 'link' },
        { sep: true }, { head: 'Header & footer' },
        { label: 'Header', cmd: 'insert-header' },
        { label: 'Footer', cmd: 'insert-footer' },
        { label: 'Page number', cmd: 'insert-page-number' },
        { sep: true },
        { label: 'Horizontal line', cmd: 'insertHorizontalRule' },
        { label: 'Page break', sc: 'Ctrl+Enter', cmd: 'insert-page-break' },
        { sep: true },
        { label: 'Special character…', dialog: 'symbol' },
        { label: 'Date', cmd: 'insert-date' },
        { label: 'Date & time', cmd: 'insert-datetime' },
        { sep: true },
        { label: 'Table of contents', cmd: 'insert-toc' }
      ] },

      { id: 'format', label: 'Format', items: [
        { label: 'Bold', sc: 'Ctrl+B', cmd: 'bold', check: function () { return qs('bold'); } },
        { label: 'Italic', sc: 'Ctrl+I', cmd: 'italic', check: function () { return qs('italic'); } },
        { label: 'Underline', sc: 'Ctrl+U', cmd: 'underline', check: function () { return qs('underline'); } },
        { label: 'Strikethrough', sc: 'Ctrl+Shift+X', cmd: 'strikeThrough', check: function () { return qs('strikeThrough'); } },
        { sep: true }, { head: 'Styles' }
      ].concat(styleItems()).concat([
        { sep: true },
        { label: 'Clear formatting', sc: 'Ctrl+Space', cmd: 'clear-formatting' },
        { label: 'Paragraph…', dialog: 'paragraph' },
        { label: 'Text colour…', dialog: 'color' },
        { label: 'Highlight colour…', dialog: 'highlight' }
      ]) },

      { id: 'tools', label: 'Tools', items: [
        { label: 'Find & Replace…', sc: 'Ctrl+F', action: function () { WW.app.openFind(false); } },
        { label: 'Word count…', dialog: 'word-count' },
        { sep: true },
        { label: 'Import HTML…', cmd: 'import-html' },
        { label: 'Sample documents…', dialog: 'samples' },
        { sep: true },
        { label: 'Spell check', check: function () { return !!U.$('#editor').spellcheck; }, action: function () { WW.app.runCmd('toggle-spellcheck'); } },
        { sep: true },
        { label: 'Preferences…', dialog: 'preferences' }
      ] },

      { id: 'help', label: 'Help', items: [
        { label: 'Keyboard shortcuts', sc: 'F1', dialog: 'shortcuts' },
        { label: 'About OpenWord', dialog: 'about' }
      ] }
    ];
  }

  /* ---------------------------------------------------------- building -- */
  function checkIcon() {
    return U.h('span', { class: 'ic', html: '<svg class="ic"><use href="#i-check"/></svg>' });
  }

  function buildItem(item) {
    if (item.sep) return U.h('li', { class: 'dropdown-sep', role: 'separator' });
    if (item.head) return U.h('li', { class: 'dropdown-head', text: item.head });
    if (item.empty) return U.h('li', { class: 'recent-empty', text: 'No recent documents' });
    var kids = [U.h('span', { class: 'label', text: item.label })];
    if (item.sc) kids.push(U.h('span', { class: 'sc', text: item.sc }));
    var li = U.h('li', { class: 'dropdown-item', role: 'menuitem' }, kids);
    li.addEventListener('click', function (e) { e.stopPropagation(); run(item); });
    li._item = item;
    return li;
  }

  function refresh(menu) {
    U.$$('.dropdown-item', menu).forEach(function (li) {
      var item = li._item;
      var old = li.querySelector('.ic');
      if (old) old.remove();
      var dis = item.disabled && item.disabled();
      if (dis) li.setAttribute('aria-disabled', 'true'); else li.removeAttribute('aria-disabled');
      if (item.check && item.check()) li.insertBefore(checkIcon(), li.firstChild);
    });
  }

  function build() {
    menubar.innerHTML = '';
    spec().forEach(function (m) {
      var menu = U.h('div', { class: 'menu', 'data-menu': m.id });
      var label = U.h('button', { class: 'menu-label', type: 'button', text: m.label });
      var dropdown = U.h('div', { class: 'dropdown', role: 'menu' },
        m.items.map(buildItem));
      label.addEventListener('click', function (e) {
        e.stopPropagation();
        if (openMenu === menu) closeAll(); else openMenuAt(menu);
      });
      menu.addEventListener('mouseenter', function () {
        if (openMenu && openMenu !== menu) openMenuAt(menu);
      });
      menu.appendChild(label);
      menu.appendChild(dropdown);
      menubar.appendChild(menu);
    });
  }

  function openMenuAt(menu) {
    closeAll();
    refresh(menu);
    menu.classList.add('open');
    openMenu = menu;
  }

  function closeAll() {
    U.$$('.menu.open', menubar).forEach(function (m) { m.classList.remove('open'); });
    openMenu = null;
  }

  /* ------------------------------------------------------ context menu -- */
  function ctxItem(item) {
    if (item.sep) return U.h('li', { class: 'ctx-sep', role: 'separator' });
    var kids = [U.h('span', { class: 'label', text: item.label })];
    if (item.sc) kids.push(U.h('span', { class: 'sc', text: item.sc }));
    var li = U.h('li', { class: 'ctx-item', role: 'menuitem' }, kids);
    if (item.disabled && item.disabled()) li.setAttribute('aria-disabled', 'true');
    else li.addEventListener('click', function () { closeCtx(); run(item); });
    return li;
  }

  function closeCtx() { if (ctxMenu) ctxMenu.hidden = true; }

  function showCtx(x, y, items) {
    ctxMenu.innerHTML = '';
    items.forEach(function (it) { ctxMenu.appendChild(ctxItem(it)); });
    ctxMenu.hidden = false;
    var r = ctxMenu.getBoundingClientRect();
    ctxMenu.style.left = Math.min(x, window.innerWidth - r.width - 6) + 'px';
    ctxMenu.style.top = Math.min(y, window.innerHeight - r.height - 6) + 'px';
  }

  function clipboardPaste() {
    if (!navigator.clipboard || !navigator.clipboard.readText) {
      WW.events.toast('Use Ctrl+Shift+V to paste plain text', 'warn');
      return;
    }
    navigator.clipboard.readText().then(function (t) {
      WW.commands.exec('paste-plain', t);
      WW.events.changed('paste');
    }).catch(function () {
      WW.events.toast('Use Ctrl+Shift+V to paste plain text', 'warn');
    });
  }

  function textCtx() {
    return [
      { label: 'Undo', sc: 'Ctrl+Z', action: function () { WW.history.undo(); }, disabled: function () { return !canUndo(); } },
      { label: 'Redo', sc: 'Ctrl+Y', action: function () { WW.history.redo(); }, disabled: function () { return !canRedo(); } },
      { sep: true },
      { label: 'Cut', sc: 'Ctrl+X', cmd: 'cut' },
      { label: 'Copy', sc: 'Ctrl+C', cmd: 'copy' },
      { label: 'Paste', sc: 'Ctrl+V', action: clipboardPaste },
      { label: 'Paste plain', sc: 'Ctrl+Shift+V', cmd: 'paste-plain' },
      { sep: true },
      { label: 'Bold', sc: 'Ctrl+B', cmd: 'bold', check: function () { return qs('bold'); } },
      { label: 'Italic', sc: 'Ctrl+I', cmd: 'italic', check: function () { return qs('italic'); } },
      { label: 'Underline', sc: 'Ctrl+U', cmd: 'underline', check: function () { return qs('underline'); } },
      { sep: true },
      { label: 'Find & Replace…', sc: 'Ctrl+F', action: function () { WW.app.openFind(false); } },
      { label: 'Paragraph…', dialog: 'paragraph' },
      { label: 'Select all', sc: 'Ctrl+A', cmd: 'select-all' }
    ];
  }

  function linkCtx(a) {
    return [
      { label: 'Open link', action: function () { window.open(a.href, '_blank', 'noopener'); } },
      { label: 'Edit link…', action: function () { WW.commands.caretNode && WW.app.runCmd('open-link-dialog', a); } },
      { label: 'Copy link', action: function () {
        if (navigator.clipboard) navigator.clipboard.writeText(a.href).then(function () { WW.events.toast('Link copied', 'ok'); });
      } },
      { sep: true },
      { label: 'Remove link', cmd: 'remove-link' }
    ];
  }

  function imageCtx(img) {
    return [
      { label: 'Cut', cmd: 'cut' },
      { label: 'Copy', cmd: 'copy' },
      { sep: true },
      { label: 'Insert link…', dialog: 'link' },
      { label: 'Reset size', action: function () {
        img.style.width = ''; img.style.height = '';
        WW.events.changed('image-resize');
        WW.pages.relayout();
      } },
      { label: 'Delete image', action: function () {
        var p = img.parentElement;
        img.remove();
        if (p && !p.textContent.trim() && !p.querySelector('img,table')) p.innerHTML = '<br>';
        WW.events.changed('delete-image');
      } }
    ];
  }

  function tableCtx() {
    return [
      { label: 'Insert row above', cmd: 'table-insert-row-above' },
      { label: 'Insert row below', cmd: 'table-insert-row-below' },
      { label: 'Insert column before', cmd: 'table-insert-col-before' },
      { label: 'Insert column after', cmd: 'table-insert-col-after' },
      { sep: true },
      { label: 'Merge cells', cmd: 'table-merge' },
      { label: 'Split cell', cmd: 'table-split' },
      { label: 'Header row', cmd: 'table-header-row', check: function () {
        var c = WW.commands.tableContext();
        return !!(c && c.row.parentElement.tagName === 'THEAD');
      } },
      { sep: true },
      { label: 'Delete row', cmd: 'table-delete-row' },
      { label: 'Delete column', cmd: 'table-delete-col' },
      { label: 'Delete table', cmd: 'table-delete' }
    ];
  }

  /* -------------------------------------------------------------- init -- */
  function init() {
    menubar = U.$('#menubar');
    ctxMenu = U.$('#ctx-menu');
    build();

    document.addEventListener('mousedown', function (e) {
      if (openMenu && !e.target.closest('.menu')) closeAll();
      if (!ctxMenu.hidden && !e.target.closest('.ctx-menu')) closeCtx();
    });
    window.addEventListener('blur', function () { closeAll(); closeCtx(); });
    document.addEventListener('scroll', closeCtx, true);

    document.addEventListener('contextmenu', function (e) {
      var editable = e.target.closest('#editor, [data-hf]');
      if (!editable) return;
      e.preventDefault();
      var img = e.target.closest('img');
      if (img) { WW.commands.selectImage(img); showCtx(e.clientX, e.clientY, imageCtx(img)); return; }
      var a = e.target.closest('a');
      if (a) { showCtx(e.clientX, e.clientY, linkCtx(a)); return; }
      if (e.target.closest('td,th')) { showCtx(e.clientX, e.clientY, tableCtx()); return; }
      showCtx(e.clientX, e.clientY, textCtx());
    });
  }

  return {
    init: init, build: build, closeAll: closeAll, closeCtx: closeCtx,
    isOpen: function () { return !!openMenu || !ctxMenu.hidden; },
    clipboardPaste: clipboardPaste
  };
})(WW.util);
