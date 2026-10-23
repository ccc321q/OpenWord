/* ==========================================================================
   OpenWord — application bootstrap and glue.
   Defines WW.APP (shared state), WW.events (changed/toast) and WW.app
   (command dispatcher, file open/save, status bar, find panel). Loaded last;
   wires every hook that index.html declares.
   ========================================================================== */
(function (U) {
  'use strict';

  var editor, header, footer, viewport;

  /* ============================================================ state == */
  var APP = WW.APP = {
    zoom: 1,
    viewMode: 'print',
    FONT_SIZES: [8, 9, 10, 11, 12, 14, 18, 24, 36, 48, 72],
    FONTS: ['Calibri', 'Carlito', 'Cambria', 'Times New Roman', 'Georgia', 'Garamond',
      'Palatino Linotype', 'Book Antiqua', 'Bookman Old Style', 'Century Schoolbook',
      'Franklin Gothic Medium', 'Segoe UI', 'Roboto', 'Open Sans', 'Lato', 'Verdana',
      'Tahoma', 'Trebuchet MS', 'Arial', 'Helvetica', 'Impact', 'Comic Sans MS',
      'Courier New', 'Consolas', 'monospace', 'serif', 'sans-serif'],
    page: {
      size: 'Letter', orientation: 'portrait',
      margins: { top: 1, right: 1, bottom: 1, left: 1, header: 0.5, footer: 0.5 }
    },
    meta: {},
    docTitle: 'Untitled document',
    dirty: false,
    canUndo: false,
    canRedo: false
  };

  /* ============================================================ events == */
  WW.events = {
    changed: function (label) {
      if (!label || label === 'typing') WW.history.pushDebounced(label || 'typing');
      else if (label !== 'undo' && label !== 'redo') WW.history.push(label);
      markDirty();
      WW.pages.relayout();
      WW.storage.scheduleAutosave();
      updateStatusSoon();
    },
    toast: function (msg, kind) {
      var wrap = U.$('#toast-wrap');
      var el = U.h('div', { class: 'toast' + (kind === 'err' || kind === 'error' ? ' err' : kind === 'ok' ? ' ok' : ''), text: msg });
      wrap.appendChild(el);
      setTimeout(function () {
        el.style.transition = 'opacity .3s';
        el.style.opacity = '0';
        setTimeout(function () { el.remove(); }, 320);
      }, kind === 'err' || kind === 'error' ? 5200 : 3200);
    }
  };

  /* ------------------------------------------------------- dirty state -- */
  function markDirty() {
    APP.dirty = true;
    setSaveState('unsaved');
  }
  function setSaveState(state) {
    var el = U.$('#save-state'), dot = U.$('#dirty-dot');
    if (state === 'unsaved') { el.textContent = 'Unsaved changes'; dot.hidden = false; }
    else if (state === 'saved') {
      var d = new Date();
      el.textContent = 'Saved ' + d.getHours() + ':' + ('0' + d.getMinutes()).slice(-2);
      dot.hidden = true; APP.dirty = false;
    } else if (state === 'error') { el.textContent = 'Autosave failed'; dot.hidden = false; }
  }

  /* ======================================================== dispatcher == */
  var APP_CMDS = {
    'new': function () { newDoc(); },
    'open': function () { U.$('#file-input').click(); },
    'save-docx': function () { saveDocx(); },
    'save-gdoc': function () { saveDocx(function () { WW.dialogs.open('gdoc'); }); },
    'save-html': function () { saveHtml(); },
    'save-txt': function () { saveTxt(); },
    'save-odt': function () { saveOdt(); },
    'print': function () { doPrint(); },
    'import-html': function () { WW.dialogs.open('import-html'); },
    'clear-autosave': function () {
      WW.storage.clearAutosave();
      WW.events.toast('Autosave copy deleted', 'ok');
    },
    'undo': function () { WW.history.undo(); },
    'redo': function () { WW.history.redo(); },
    'cut': function () { document.execCommand('cut'); },
    'copy': function () { document.execCommand('copy'); },
    'paste': function () { WW.menus.clipboardPaste(); },
    'toggle-ruler': function () {
      var on = U.$('#ruler-wrap').hidden;
      WW.pages.setRuler(on);
      syncToolbar();
    },
    'toggle-marks': function () {
      WW.pages.setMarks(!document.body.classList.contains('show-marks'));
      syncToolbar();
    },
    'toggle-grid': function () {
      WW.pages.setGuides(!document.body.classList.contains('show-guides'));
      syncToolbar();
    },
    'toggle-spellcheck': function () {
      editor.spellcheck = !editor.spellcheck;
      WW.storage.setPref('spellcheck', editor.spellcheck);
      syncToolbar();
    },
    'toggle-theme': function () {
      var dark = !document.body.classList.contains('theme-dark');
      document.body.classList.toggle('theme-dark', dark);
      document.body.classList.toggle('theme-light', !dark);
      WW.storage.setPref('theme', dark ? 'dark' : 'light');
      WW.pages.drawRuler();
    },
    'zoom-in': function () { WW.pages.setZoom(APP.zoom + 0.1); },
    'zoom-out': function () { WW.pages.setZoom(APP.zoom - 0.1); },
    'zoom-reset': function () { WW.pages.setZoom(1); },
    'orientation': function () {
      APP.page.orientation = APP.page.orientation === 'portrait' ? 'landscape' : 'portrait';
      WW.pages.applyPage();
      WW.events.changed('orientation');
    },
    'insert-header': function () { focusHf(header); },
    'insert-footer': function () { focusHf(footer); },
    'open-link-dialog': function (a) {
      if (a) {
        var r = document.createRange();
        r.selectNodeContents(a);
        var s = window.getSelection();
        s.removeAllRanges(); s.addRange(r);
      }
      WW.dialogs.open('link');
    }
  };

  function runCmd(cmd, value) {
    if (APP_CMDS[cmd]) { APP_CMDS[cmd](value); syncToolbar(); return; }
    var ok = WW.commands.exec(cmd, value);
    if (ok) WW.events.changed(cmd);
    syncToolbar();
  }

  function focusHf(el) {
    if (!el.innerHTML.trim()) el.innerHTML = '<p><br></p>';
    el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    el.focus();
    updateHfEmpty();
  }

  /* ============================================================ chrome == */
  function wireChrome() {
    // ribbon tabs
    U.$$('.ribbon-tab').forEach(function (tab) {
      tab.addEventListener('click', function () {
        U.$$('.ribbon-tab').forEach(function (t) { t.classList.remove('active'); });
        U.$$('.ribbon-panel').forEach(function (p) { p.classList.remove('active'); });
        tab.classList.add('active');
        U.$('.ribbon-panel[data-panel="' + tab.dataset.tab + '"]').classList.add('active');
      });
    });

    // generic button hooks
    document.addEventListener('click', function (e) {
      var el = e.target.closest('[data-cmd],[data-dialog],[data-panel-find]');
      if (!el || el.tagName === 'SELECT' || el.disabled) return;
      if (el.dataset.dialog) { WW.dialogs.open(el.dataset.dialog); return; }
      if (el.hasAttribute('data-panel-find')) { openFind(false); return; }
      runCmd(el.dataset.cmd);
    });

    // split colour buttons: apply last colour
    U.$('#color-main').addEventListener('click', function () {
      WW.commands.exec('foreColor'); WW.events.changed('color');
    });
    U.$('#highlight-main').addEventListener('click', function () {
      WW.commands.exec('hiliteColor'); WW.events.changed('highlight');
    });

    // selects
    populateSelects();
    U.$('#font-family').addEventListener('change', function () {
      WW.commands.exec('font-family', this.value); WW.events.changed('font-family');
    });
    U.$('#font-size').addEventListener('change', function () {
      WW.commands.exec('font-size', this.value); WW.events.changed('font-size');
    });
    U.$('#block-style').addEventListener('change', function () {
      WW.commands.exec('formatBlock', this.value); WW.events.changed('block-style');
    });
    U.$('#line-spacing').addEventListener('change', function () {
      if (this.value === 'dialog') { WW.dialogs.open('paragraph'); }
      else if (this.value) { WW.commands.exec('setLineSpacing', this.value); WW.events.changed('line-spacing'); }
      this.value = '';
    });
    U.$('#page-size').addEventListener('change', function () {
      APP.page.size = this.value;
      WW.pages.applyPage();
      WW.events.changed('page-size');
    });
    U.$('#margins-preset').addEventListener('change', function () {
      if (this.value === 'dialog') { WW.dialogs.open('page-setup'); }
      else {
        var v = parseFloat(this.value);
        WW.dialogs.applyPageSetup({
          size: APP.page.size, orientation: APP.page.orientation,
          margins: { top: v, right: v, bottom: v, left: v, header: 0.5, footer: 0.5 }
        });
      }
      this.value = '';
    });
    U.$('#view-mode').addEventListener('change', function () { WW.pages.setViewMode(this.value); });
    U.$('#status-view').addEventListener('change', function () { WW.pages.setViewMode(this.value); });
    U.$('#zoom-range').addEventListener('input', function () { WW.pages.setZoom(this.value / 100, true); });

    // theme + find panel buttons
    U.$('#btn-theme').addEventListener('click', function () { runCmd('toggle-theme'); });
    U.$('#find-close').addEventListener('click', closeFind);
    U.$('#find-input').addEventListener('input', runSearch);
    U.$('#find-input').addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); e.shiftKey ? WW.find.prev() : WW.find.next(); }
      if (e.key === 'Escape') closeFind();
    });
    U.$('#find-case').addEventListener('change', runSearch);
    U.$('#find-word').addEventListener('change', runSearch);
    U.$('#find-next').addEventListener('click', function () { WW.find.next(); });
    U.$('#find-prev').addEventListener('click', function () { WW.find.prev(); });
    U.$('#replace-one').addEventListener('click', function () {
      WW.find.replaceOne(findTerm(), U.$('#replace-input').value);
      WW.events.changed('replace'); runSearch();
    });
    U.$('#replace-all').addEventListener('click', function () {
      var n = WW.find.replaceAll(findTerm(), U.$('#replace-input').value);
      WW.events.toast(n + (n === 1 ? ' replacement' : ' replacements'), 'ok');
      WW.events.changed('replace-all'); runSearch();
    });
    WW.find.onChange(function (stats) {
      var el = U.$('#find-status');
      if (!stats || !stats.total) { el.textContent = 'No results'; el.classList.add('bad'); return; }
      el.classList.remove('bad');
      el.textContent = (stats.index + 1) + ' of ' + stats.total;
    });

    // document title rename
    U.$('#doc-title').addEventListener('click', renameTitle);

    // file input
    U.$('#file-input').addEventListener('change', function () {
      var files = Array.prototype.slice.call(this.files);
      this.value = '';
      openFiles(files);
    });

    // editor events
    [editor, header, footer].forEach(function (el) {
      el.addEventListener('input', function () {
        WW.events.changed('typing');
        updateHfEmpty();
      });
      el.addEventListener('paste', onPaste);
    });
    document.addEventListener('selectionchange', U.throttle(function () {
      WW.pages.markFocusBlock();
      syncToolbar();
    }, 120));

    // drag & drop
    var hint = U.h('div', { class: 'dropzone-hint' }, ['Drop a document to open it']);
    document.body.appendChild(hint);
    window.addEventListener('dragover', function (e) {
      e.preventDefault();
      document.body.classList.add('dropping');
    });
    window.addEventListener('dragleave', function (e) {
      if (e.target === document.documentElement || (!e.relatedTarget)) document.body.classList.remove('dropping');
    });
    window.addEventListener('drop', function (e) {
      e.preventDefault();
      document.body.classList.remove('dropping');
      if (e.dataTransfer.files && e.dataTransfer.files.length) {
        openFiles(Array.prototype.slice.call(e.dataTransfer.files));
        return;
      }
      var html = e.dataTransfer.getData('text/html');
      var text = e.dataTransfer.getData('text/plain');
      if (html) {
        WW.commands.exec('insert-html', WW.htmlio.fragmentFromHtml(html).innerHTML);
        WW.events.changed('drop-html');
      } else if (text) {
        WW.commands.exec('insert-text', text);
        WW.events.changed('drop-text');
      }
    });

    // history state -> buttons
    WW.history.onChange(function (canUndo, canRedo) {
      APP.canUndo = canUndo; APP.canRedo = canRedo;
      U.$('#qat-undo').disabled = !canUndo;
      U.$('#qat-redo').disabled = !canRedo;
    });

    // warn on close with unsaved work
    window.addEventListener('beforeunload', function (e) {
      if (APP.dirty) { e.preventDefault(); e.returnValue = ''; }
    });
  }

  /* --------------------------------------------------------- selects -- */
  function populateSelects() {
    var ff = U.$('#font-family');
    ff.innerHTML = '';
    ff.appendChild(U.h('option', { value: '', text: 'Font' }));
    APP.FONTS.forEach(function (f) {
      ff.appendChild(U.h('option', { value: f, text: f, style: 'font-family:' + f }));
    });
    var fs = U.$('#font-size');
    fs.innerHTML = '';
    APP.FONT_SIZES.forEach(function (s) { fs.appendChild(U.h('option', { value: String(s), text: String(s) })); });

    var bs = U.$('#block-style');
    bs.innerHTML = '';
    [['Normal', 'Normal'], ['Heading1', 'Heading 1'], ['Heading2', 'Heading 2'], ['Heading3', 'Heading 3'],
      ['Heading4', 'Heading 4'], ['Heading5', 'Heading 5'], ['Heading6', 'Heading 6'],
      ['Title', 'Title'], ['Subtitle', 'Subtitle'], ['Quote', 'Quote'],
      ['IntenseQuote', 'Intense Quote'], ['Code', 'Code']].forEach(function (o) {
        bs.appendChild(U.h('option', { value: o[0], text: o[1] }));
      });

    // style gallery
    var gallery = U.$('#style-gallery');
    gallery.innerHTML = '';
    [['Normal', 'Normal', ''], ['Heading1', 'Heading 1', 'font-size:13pt;font-weight:600;color:#2b579a'],
      ['Heading2', 'Heading 2', 'font-size:12pt;font-weight:600;color:#2b579a'],
      ['Heading3', 'Heading 3', 'font-size:11.5pt;font-weight:600;color:#2b579a'],
      ['Title', 'Title', 'font-size:15pt'], ['Subtitle', 'Subtitle', 'font-size:11pt;color:#666'],
      ['Quote', 'Quote', 'font-style:italic'], ['IntenseQuote', 'Intense Quote', 'font-style:italic;color:#2b579a'],
      ['Code', 'Code', 'font-family:Consolas,monospace;font-size:9.5pt']].forEach(function (s) {
      var chip = U.h('button', {
        class: 'style-chip', type: 'button', title: 'Apply “' + s[1] + '”',
        onclick: function () {
          WW.commands.exec('formatBlock', s[0]);
          WW.events.changed('block-style');
          syncToolbar();
        }
      }, [U.h('span', { text: s[1], style: s[2] })]);
      chip.dataset.style = s[0];
      gallery.appendChild(chip);
    });
  }

  /* ------------------------------------------------------ toolbar sync -- */
  function syncToolbar() {
    var C = WW.commands;
    // toggle buttons driven by editor state
    var toggles = {
      'bold': 'bold', 'italic': 'italic', 'underline': 'underline', 'strikeThrough': 'strikeThrough',
      'subscript': 'subscript', 'superscript': 'superscript',
      'insertUnorderedList': 'insertUnorderedList', 'insertOrderedList': 'insertOrderedList',
      'justifyLeft': 'justifyLeft', 'justifyCenter': 'justifyCenter',
      'justifyRight': 'justifyRight', 'justifyFull': 'justifyFull'
    };
    Object.keys(toggles).forEach(function (cmd) {
      U.$$('button[data-cmd="' + cmd + '"]').forEach(function (b) {
        b.classList.toggle('on', !!C.queryState(cmd));
      });
    });
    // view toggles
    U.$('#btn-ruler').classList.toggle('on', !U.$('#ruler-wrap').hidden);
    U.$('#btn-marks').classList.toggle('on', document.body.classList.contains('show-marks'));
    U.$('#btn-grid').classList.toggle('on', document.body.classList.contains('show-guides'));
    U.$('#btn-spellcheck').classList.toggle('on', !!editor.spellcheck);
    U.$('#btn-painter').classList.toggle('on', !!C.state.painter);

    // selects (skip the one being interacted with)
    var act = document.activeElement;
    var fam = C.queryValue('font-family');
    var famSel = U.$('#font-family');
    if (act !== famSel) {
      if (!U.$$('#font-family option').some(function (o) { return o.value === fam; })) {
        var opt = U.h('option', { value: fam, text: fam });
        famSel.appendChild(opt);
      }
      famSel.value = fam;
    }
    var sizeSel = U.$('#font-size');
    if (act !== sizeSel) {
      var size = C.queryValue('font-size');
      sizeSel.value = size;
      if (sizeSel.value !== size) {
        var o2 = U.h('option', { value: size, text: size });
        sizeSel.appendChild(o2);
        sizeSel.value = size;
      }
    }
    var bsSel = U.$('#block-style');
    if (act !== bsSel) bsSel.value = C.queryValue('block-style') || 'Normal';
    var lsSel = U.$('#line-spacing');
    if (act !== lsSel && lsSel.value !== 'dialog') lsSel.value = C.queryValue('line-spacing') || '';

    // style gallery highlight
    var cur = C.currentBlockStyle();
    U.$$('.style-chip').forEach(function (c) { c.classList.toggle('on', c.dataset.style === cur); });

    // status format readout
    var flags = [];
    if (C.queryState('bold')) flags.push('B');
    if (C.queryState('italic')) flags.push('I');
    if (C.queryState('underline')) flags.push('U');
    if (C.queryState('strikeThrough')) flags.push('S');
    U.$('#status-format').textContent = (fam || '') + ' ' + (sizeSel.value || '') + (flags.length ? '  ' + flags.join(' ') : '');

    // page readout
    U.$('#status-page').textContent = 'Page ' + WW.pages.pageAtCaret() + ' of ' + WW.pages.totalPages();
  }

  /* ------------------------------------------------------ status counts -- */
  var statusPending = false;
  function updateStatusSoon() {
    if (statusPending) return;
    statusPending = true;
    requestAnimationFrame(function () {
      statusPending = false;
      updateStatus();
    });
  }
  function updateStatus() {
    var text = (editor.textContent || '').replace(/[\u200b\ufeff]/g, '');
    var words = (text.match(/[^\s]+/g) || []).length;
    U.$('#status-words').textContent = U.pluralize(words, 'word');
    U.$('#status-chars').textContent = U.pluralize(text.replace(/\s/g, '').length, 'character');
    U.$('#status-read').textContent = Math.max(text.trim() ? 1 : 0, Math.ceil(words / 200)) + ' min read';
    U.$('#status-page').textContent = 'Page ' + WW.pages.pageAtCaret() + ' of ' + WW.pages.totalPages();
  }

  function updateHfEmpty() {
    [[header, 'Header'], [footer, 'Footer']].forEach(function (pair) {
      var empty = !U.stripZwsp(pair[0].textContent || '').trim() && !pair[0].querySelector('img,.ww-field');
      pair[0].classList.toggle('empty', empty);
      pair[0].dataset.placeholder = pair[1];
    });
  }

  /* ============================================================ paste == */
  function onPaste(e) {
    var dt = e.clipboardData || window.clipboardData;
    if (!dt) return;
    e.preventDefault();
    var html = dt.getData('text/html');
    var text = dt.getData('text/plain');
    if (html) {
      var host = WW.htmlio.fragmentFromHtml(html);
      WW.commands.exec('insert-html', host.innerHTML);
    } else if (text) {
      WW.commands.exec('insert-text', text);
    }
    WW.events.changed('paste');
  }

  /* ============================================================= find == */
  function findTerm() { return U.$('#find-input').value; }
  function runSearch() {
    WW.find.search(findTerm(), {
      caseSensitive: U.$('#find-case').checked,
      wholeWord: U.$('#find-word').checked
    });
  }
  function openFind(showReplace) {
    var panel = U.$('#find-panel');
    panel.hidden = false;
    var input = showReplace ? U.$('#replace-input') : U.$('#find-input');
    input.focus();
    input.select();
    runSearch();
  }
  function closeFind() {
    U.$('#find-panel').hidden = true;
    WW.find.clear();
    editor.focus();
  }

  /* ============================================================ title == */
  function renameTitle() {
    var btn = U.$('#doc-title');
    var wrap = btn.parentElement;
    var inp = U.h('input', {
      type: 'text', value: APP.docTitle, class: 'doc-title',
      style: 'width:220px;background:var(--bg-chrome);border:1px solid var(--blue);border-radius:3px;padding:2px 6px;font:inherit;color:inherit'
    });
    wrap.replaceChild(inp, btn);
    inp.focus();
    inp.select();
    var done = false;
    function commit(save) {
      if (done) return;
      done = true;
      if (save && inp.value.trim()) setDocTitle(inp.value.trim());
      wrap.replaceChild(btn, inp);
    }
    inp.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') commit(true);
      if (e.key === 'Escape') commit(false);
    });
    inp.addEventListener('blur', function () { commit(true); });
  }
  function setDocTitle(name) {
    APP.docTitle = name || 'Untitled document';
    U.$('#doc-title').textContent = APP.docTitle;
    document.title = APP.docTitle + ' — OpenWord';
    WW.pages.updateFields();
  }

  /* ========================================================= documents == */
  function cleanHtml(el) {
    var clone = el.cloneNode(true);
    U.$$('.ww-img-resizer', clone).forEach(function (n) { n.remove(); });
    U.$$('.ww-selected', clone).forEach(function (n) { n.classList.remove('ww-selected'); });
    U.$$('.ww-focus', clone).forEach(function (n) { n.classList.remove('ww-focus'); });
    return clone.innerHTML;
  }

  function collectDoc() {
    return {
      title: APP.docTitle,
      html: cleanHtml(editor),
      headerHtml: cleanHtml(header),
      footerHtml: cleanHtml(footer),
      page: U.deepClone(APP.page),
      meta: U.deepClone(APP.meta)
    };
  }

  function loadDoc(doc) {
    editor.innerHTML = doc.html || '<p><br></p>';
    header.innerHTML = doc.headerHtml || '';
    footer.innerHTML = doc.footerHtml || '';
    APP.page = doc.page || U.deepClone(APP.page);
    APP.meta = doc.meta || {};
    setDocTitle(doc.title || 'Untitled document');
    WW.pages.applyPage();
    WW.history.reset(doc.label || 'open');
    APP.dirty = false;
    setSaveState('saved');
    updateHfEmpty();
    WW.pages.updateFields();
    updateStatus();
    syncToolbar();
  }

  function newDoc() {
    function go() {
      WW.storage.clearAutosave();
      loadDoc({ html: '<p><br></p>', title: 'Untitled document', page: U.deepClone(APP.page), fresh: true });
      editor.focus();
    }
    if (APP.dirty) {
      WW.dialogs.confirm('New document', 'Discard unsaved changes to “' + APP.docTitle + '”?', 'Discard', true)
        .then(function (yes) { if (yes) go(); });
    } else go();
  }

  /* ------------------------------------------------------------- open -- */
  function openFiles(files) {
    if (!files.length) return;
    var file = files[0];
    if (files.length > 1) WW.events.toast('Opening “' + file.name + '” (first file only)', 'warn');
    var ext = U.extensionOf(file.name);

    function applyImported(res, name) {
      function go() {
        loadDoc({
          html: res.html, headerHtml: res.headerHtml || res.header || '',
          footerHtml: res.footerHtml || res.footer || '',
          page: res.page || APP.page, meta: res.meta || {},
          title: U.baseNameOf(name)
        });
        WW.storage.addRecent(U.baseNameOf(name));
        if (res.warnings && res.warnings.length) {
          WW.events.toast(res.warnings.length === 1 ? res.warnings[0] :
            res.warnings[0] + ' (+' + (res.warnings.length - 1) + ' more)', 'warn');
        }
      }
      if (APP.dirty) {
        WW.dialogs.confirm('Open document', 'Discard unsaved changes to “' + APP.docTitle + '”?', 'Open', true)
          .then(function (yes) { if (yes) go(); });
      } else go();
    }

    var busy = splash('Opening ' + file.name + '…');
    var job;
    if (ext === 'docx' || ext === 'dotx') {
      job = U.readFileAsArrayBuffer(file).then(function (buf) { return WW.docxImport.parseFile(buf); });
    } else if (ext === 'odt' || ext === 'ott') {
      job = U.readFileAsArrayBuffer(file).then(function (buf) { return WW.odt.parseOdt(buf); });
    } else if (ext === 'html' || ext === 'htm' || ext === 'gdoc') {
      job = U.readFileAsText(file).then(function (text) {
        return { html: WW.htmlio.fragmentFromHtml(text).innerHTML };
      });
    } else if (ext === 'txt') {
      job = U.readFileAsText(file).then(function (text) {
        return { html: WW.htmlio.fromPlainText(text).innerHTML };
      });
    } else if (ext === 'rtf') {
      job = U.readFileAsText(file).then(function (text) {
        return { html: WW.htmlio.fromPlainText(rtfToText(text)).innerHTML };
      });
    } else {
      busy.remove();
      WW.events.toast('Unsupported file type: .' + ext, 'err');
      return;
    }
    job.then(function (res) { busy.remove(); applyImported(res, file.name); })
      .catch(function (err) {
        busy.remove();
        WW.events.toast('Could not open ' + file.name + ': ' + (err && err.message || err), 'err');
      });
  }

  /** Minimal RTF → plain text: enough for pasted notes, not a full parser. */
  function rtfToText(rtf) {
    var text = String(rtf);
    text = text.replace(/\\'([0-9a-fA-F]{2})/g, function (m, h) { return String.fromCharCode(parseInt(h, 16)); });
    text = text.replace(/\\u(-?\d+)\s?/g, function (m, n) { var c = parseInt(n, 10); return c < 0 ? String.fromCharCode(c + 65536) : String.fromCharCode(c); });
    text = text.replace(/\\par[d]?\b/g, '\n').replace(/\\line\b/g, '\n').replace(/\\tab\b/g, '\t');
    text = text.replace(/\\(fonttbl|colortbl|stylesheet|info|pict|object)[\s\S]*?(?=\})/g, '');
    text = text.replace(/\\[a-zA-Z]{1,32}-?\d{0,10} ?/g, '');
    text = text.replace(/[{}]/g, '');
    text = text.replace(/\r/g, '');
    return text.replace(/\n{3,}/g, '\n\n').trim();
  }

  function splash(text) {
    var el = U.h('div', { class: 'splash-loading' }, [U.h('div', { text: text })]);
    document.body.appendChild(el);
    return el;
  }

  /* ------------------------------------------------------------- save -- */
  function finishSave(name, blob, warnings, after) {
    U.downloadBlob(blob, name);
    APP.dirty = false;
    setSaveState('saved');
    WW.storage.addRecent(APP.docTitle);
    if (warnings && warnings.length) {
      WW.events.toast(warnings.length === 1 ? warnings[0] :
        warnings[0] + ' (+' + (warnings.length - 1) + ' more)', 'warn');
    } else {
      WW.events.toast('Saved ' + name, 'ok');
    }
    if (after) after();
  }

  function saveDocx(after) {
    var busy = splash('Exporting .docx…');
    WW.docxExport.exportDocument({
      editor: editor, header: header, footer: footer,
      page: APP.page, meta: Object.assign({ title: APP.docTitle }, APP.meta)
    }).then(function (res) {
      busy.remove();
      finishSave(APP.docTitle + '.docx', res.blob, res.warnings, after);
    }).catch(function (err) {
      busy.remove();
      WW.events.toast('Export failed: ' + (err && err.message || err), 'err');
    });
  }

  function saveOdt() {
    var busy = splash('Exporting .odt…');
    WW.odt.exportOdt({
      editor: editor, header: header, footer: footer,
      page: APP.page, meta: Object.assign({ title: APP.docTitle }, APP.meta)
    }).then(function (res) {
      busy.remove();
      finishSave(APP.docTitle + '.odt', res.blob, res.warnings);
    }).catch(function (err) {
      busy.remove();
      WW.events.toast('Export failed: ' + (err && err.message || err), 'err');
    });
  }

  function saveHtml() {
    var doc = WW.htmlio.toHtmlDocument(editor, {
      title: APP.docTitle, header: header, footer: footer, page: APP.page
    });
    finishSave(APP.docTitle + '.html', new Blob([doc], { type: 'text/html;charset=utf-8' }));
  }

  function saveTxt() {
    var text = WW.htmlio.toPlainText(editor);
    finishSave(APP.docTitle + '.txt', new Blob([text], { type: 'text/plain;charset=utf-8' }));
  }

  function doPrint() {
    var st = U.$('#print-page-style');
    if (!st) {
      st = U.h('style', { id: 'print-page-style' });
      document.head.appendChild(st);
    }
    var w = WW.pages.pageWidthPx() / 96, h = WW.pages.pageHeightPx() / 96;
    st.textContent = '@page { size: ' + w.toFixed(2) + 'in ' + h.toFixed(2) + 'in; margin: 0; }';
    WW.pages.updateFields();
    setTimeout(function () { window.print(); }, 60);
  }

  /* ------------------------------------------------------- recent open -- */
  function openSampleByName(name) {
    var s = WW.dialogs.SAMPLES.filter(function (x) { return x.name === name; })[0];
    if (!s) { WW.events.toast('That document is not available as a sample', 'warn'); return; }
    function go() { loadDoc({ html: s.html, title: s.name }); }
    if (APP.dirty) {
      WW.dialogs.confirm('Open document', 'Discard unsaved changes to “' + APP.docTitle + '”?', 'Open', true)
        .then(function (yes) { if (yes) go(); });
    } else go();
  }

  /* ============================================================ prefs == */
  function applyPrefs() {
    var p = WW.storage.getPrefs();
    document.body.classList.toggle('theme-dark', p.theme === 'dark');
    document.body.classList.toggle('theme-light', p.theme !== 'dark');
    WW.pages.setZoom(p.zoom || 1);
    WW.pages.setViewMode(p.view || 'print');
    WW.pages.setRuler(p.ruler !== false);
    WW.pages.setMarks(!!p.marks);
    WW.pages.setGuides(p.guides !== false);
    editor.spellcheck = p.spellcheck !== false;
    editor.style.fontFamily = p.defaultFont ? '"' + p.defaultFont + '", ' + editor.style.fontFamily : '';
    if (p.defaultSize) editor.style.fontSize = p.defaultSize + 'pt';
    WW.commands.state.defaultFont = p.defaultFont || 'Calibri';
    WW.commands.state.defaultSize = p.defaultSize || 11;
  }

  /* ============================================================== boot == */
  function boot() {
    editor = U.$('#editor');
    header = U.$('#page-header');
    footer = U.$('#page-footer');
    viewport = U.$('#viewport');

    WW.storage.loadPrefs();
    if (!editor.innerHTML.trim()) editor.innerHTML = '<p><br></p>';
    WW.commands.init();
    WW.history.init();
    WW.pages.init();
    WW.dialogs.init();
    WW.menus.init();
    WW.shortcuts.init();
    WW.commands.attachImageUI();

    wireChrome();
    applyPrefs();

    updateHfEmpty();
    updateStatus();
    syncToolbar();
    setSaveState('saved');

    var auto = WW.storage.loadAutosave();
    if (auto) {
      WW.dialogs.restoreDialog(auto).then(function (yes) {
        if (yes) {
          loadDoc({
            html: auto.html, headerHtml: auto.headerHtml, footerHtml: auto.footerHtml,
            page: auto.page, meta: auto.meta, title: auto.title
          });
        } else {
          WW.storage.clearAutosave();
        }
        editor.focus();
      });
    }
    WW.pages.drawRuler();
  }

  WW.app = {
    runCmd: runCmd, collectDoc: collectDoc, loadDoc: loadDoc, newDoc: newDoc,
    openFiles: openFiles, openFind: openFind, closeFind: closeFind,
    updateStatus: updateStatus, syncToolbar: syncToolbar, setSaveState: setSaveState,
    setDocTitle: setDocTitle, applyPrefs: applyPrefs, openSampleByName: openSampleByName,
    escape: function () {
      if (!U.$('#ctx-menu').hidden) { WW.menus.closeCtx(); return; }
      if (WW.dialogs.isOpen()) { WW.dialogs.close(); return; }
      if (WW.menus.isOpen()) { WW.menus.closeAll(); return; }
      if (!U.$('#find-panel').hidden) { closeFind(); return; }
      if (document.activeElement) document.activeElement.blur();
    }
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})(WW.util);
