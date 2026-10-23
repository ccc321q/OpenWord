/* ==========================================================================
   OpenWord — dialogs: a small modal engine plus every dialog the chrome uses.
   Exposes WW.dialogs. Dialogs build their DOM with U.h so no innerHTML of
   user data is ever needed here.
   ========================================================================== */
WW.dialogs = (function (U) {
  'use strict';

  var backdrop, modal, titleEl, bodyEl, footEl, onClose = null;

  /* ============================================================ engine == */
  function show(opts) {
    titleEl.textContent = opts.title || '';
    bodyEl.innerHTML = '';
    footEl.innerHTML = '';
    if (opts.wide) modal.classList.add('wide'); else modal.classList.remove('wide');
    if (typeof opts.body === 'string') bodyEl.innerHTML = opts.body;
    else if (opts.body) bodyEl.appendChild(opts.body);
    (opts.buttons || []).forEach(function (b) { footEl.appendChild(b); });
    if (!footEl.children.length) footEl.style.display = 'none';
    else footEl.style.display = '';
    onClose = opts.onClose || null;
    backdrop.hidden = false;
    var first = bodyEl.querySelector('input,select,textarea,button');
    if (first) setTimeout(function () { first.focus(); if (first.select) first.select(); }, 30);
    return { close: close };
  }

  function close() {
    if (backdrop.hidden) return;
    backdrop.hidden = true;
    var cb = onClose; onClose = null;
    if (cb) { try { cb(); } catch (e) { /* ignore */ } }
  }
  function isOpen() { return !backdrop.hidden; }

  function btn(label, cls, onClick) {
    return U.h('button', { class: 'btn' + (cls ? ' ' + cls : ''), type: 'button', text: label, onclick: onClick });
  }

  function confirm(title, message, okLabel, danger) {
    return new Promise(function (resolve) {
      var done = false;
      function finish(v) { if (!done) { done = true; resolve(v); } }
      show({
        title: title,
        body: U.h('div', null, [U.h('p', { text: message })]),
        buttons: [
          btn('Cancel', '', function () { finish(false); close(); }),
          btn(okLabel || 'OK', danger ? 'danger' : 'primary', function () { finish(true); close(); })
        ],
        onClose: function () { finish(false); }
      });
    });
  }

  function alert(title, messageHtml) {
    show({
      title: title,
      body: U.h('div', null, [U.h('div', { html: messageHtml })]),
      buttons: [btn('OK', 'primary', close)]
    });
  }

  function field(labelText, input) {
    return U.h('div', { class: 'field' }, [U.h('label', { text: labelText }), input]);
  }
  function textInput(value, type) {
    return U.h('input', { type: type || 'text', value: value == null ? '' : String(value) });
  }
  function selectInput(options, value) {
    var s = U.h('select', null, options.map(function (o) {
      return U.h('option', { value: o[0], text: o[1] });
    }));
    if (value != null) s.value = String(value);
    return s;
  }

  /* =========================================================== colours == */
  var THEME_COLORS = ['#2b579a', '#5b7a9d', '#1f4377', '#0563c1', '#2b579a', '#e8eef8', '#f5f5f5', '#8f8f8f', '#444444', '#000000'];
  var STD_COLORS = ['#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'];
  var HL_COLORS = ['#ffff00', '#ffff99', '#99ff66', '#ccffff', '#ff99cc', '#ffd966', '#ff8c00', '#c0c0c0', '#808080', '#00ff00'];

  function paletteGrid(colors, onPick) {
    var grid = U.h('div', { class: 'palette' }, colors.map(function (c) {
      return U.h('button', {
        class: 'sw', type: 'button', title: c,
        style: 'background:' + c,
        onclick: function () { onPick(c); close(); }
      });
    }));
    return grid;
  }

  function colorDialog(isHighlight) {
    var custom = U.h('input', { type: 'color', value: isHighlight ? '#ffff00' : '#c00000' });
    var body = U.h('div', null, [
      U.h('div', { class: 'palette-title', text: isHighlight ? 'Text highlight colour' : 'Theme colours' }),
      paletteGrid(isHighlight ? HL_COLORS : THEME_COLORS, function (c) { apply(c); }),
      U.h('div', { class: 'palette-title', text: 'Standard colours' }),
      paletteGrid(STD_COLORS, function (c) { apply(c); }),
      U.h('div', { class: 'custom-color-row' }, [
        custom,
        btn('Apply custom', 'primary', function () { apply(custom.value.toUpperCase()); close(); })
      ]),
      U.h('p', { class: 'hint', text: 'Tip: the main colour button re-applies your last colour.' })
    ]);
    function apply(c) {
      WW.commands.exec(isHighlight ? 'hiliteColor' : 'foreColor', c);
      WW.events.changed(isHighlight ? 'highlight' : 'color');
    }
    show({ title: isHighlight ? 'Text highlight colour' : 'Text colour', body: body, buttons: [btn('Cancel', '', close)] });
  }

  /* ============================================================= image == */
  function imageDialog() {
    var urlIn = textInput('', 'text');
    urlIn.placeholder = 'https://example.com/photo.jpg';
    var status = U.h('span', { class: 'hint', text: 'Images are embedded so the document stays self-contained.' });

    function insertDataUrl(dataUrl) {
      WW.commands.exec('insert-image', { dataUrl: dataUrl });
      WW.events.changed('insert-image');
      close();
    }
    function fromUrl() {
      var url = urlIn.value.trim();
      if (!url) return;
      status.textContent = 'Loading…';
      var img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = function () {
        try {
          var cv = document.createElement('canvas');
          cv.width = img.naturalWidth; cv.height = img.naturalHeight;
          cv.getContext('2d').drawImage(img, 0, 0);
          insertDataUrl(cv.toDataURL('image/png'));
        } catch (e) {
          // canvas tainted by a server without CORS — embed by reference instead
          insertDataUrl(url);
        }
      };
      img.onerror = function () { status.textContent = 'Could not load that URL.'; };
      img.src = url;
    }

    var fileBtn = btn('Choose file…', '', function () {
      var inp = U.$('#image-input');
      inp.onchange = function () {
        var f = inp.files && inp.files[0];
        inp.value = ''; inp.onchange = null;
        if (f) U.readFileAsDataURL(f).then(insertDataUrl);
      };
      inp.click();
    });

    var tabs = miniTabs([
      ['File', U.h('div', null, [fileBtn, U.h('p', { class: 'hint', text: 'PNG, JPEG, GIF, WebP or SVG.' })])],
      ['From URL', U.h('div', null, [
        field('URL', urlIn),
        btn('Load', 'primary', fromUrl), U.h('div', null, [status])
      ])]
    ]);
    show({ title: 'Insert picture', body: tabs, buttons: [btn('Cancel', '', close)] });
  }

  /* ============================================================== link == */
  function linkDialog() {
    var sel = window.getSelection();
    var selectedText = sel && !sel.isCollapsed ? sel.toString() : '';
    var existing = WW.commands.caretNode() && WW.commands.caretNode().closest &&
      WW.commands.caretNode().closest('a');
    var textIn = textInput(existing ? (selectedText || existing.textContent) : selectedText);
    var urlIn = textInput(existing ? existing.getAttribute('href') : 'https://');
    var kids = [field('Text to display', textIn), field('URL', urlIn)];
    if (existing) kids.push(U.h('div', { class: 'field' }, [U.h('span'), btn('Remove link', 'danger', function () {
      WW.commands.exec('remove-link'); WW.events.changed('remove-link'); close();
    })]));
    var body = U.h('div', null, kids);
    show({
      title: 'Insert link', body: body,
      buttons: [
        btn('Cancel', '', close),
        btn('Insert', 'primary', function () {
          var url = urlIn.value.trim();
          if (!url) { urlIn.focus(); return; }
          if (!/^(https?:|mailto:|#|\/|\.)/i.test(url)) url = 'https://' + url;
          WW.commands.exec('insert-link', { url: url, text: textIn.value || url });
          WW.events.changed('insert-link');
          close();
        })
      ]
    });
  }

  /* ============================================================= table == */
  function tableDialog() {
    var COLS = 10, ROWS = 8, rows = 2, cols = 3, headerRow = true;
    var preview = U.h('div', { class: 'tp-preview', text: '3 × 2 table' });
    var cells = [];
    function paint() {
      cells.forEach(function (c) {
        c.el.classList.toggle('on', c.r < rows && c.c < cols);
      });
      preview.textContent = cols + ' × ' + rows + ' table' + (headerRow ? ' with header row' : '');
    }
    var grid = U.h('div', { class: 'table-picker' });
    for (var r = 0; r < ROWS; r++) {
      for (var c = 0; c < COLS; c++) {
        (function (r, c) {
          var el = U.h('div', { class: 'tp-cell', onmouseenter: function () { rows = r + 1; cols = c + 1; paint(); } });
          cells.push({ el: el, r: r, c: c });
          grid.appendChild(el);
        })(r, c);
      }
    }
    var hdr = U.h('input', { type: 'checkbox', checked: 'checked' });
    hdr.onchange = function () { headerRow = hdr.checked; paint(); };
    var body = U.h('div', null, [
      grid, preview,
      U.h('label', { class: 'chk', style: 'margin-top:10px' }, [hdr, ' Header row'])
    ]);
    show({
      title: 'Insert table', body: body,
      buttons: [
        btn('Cancel', '', close),
        btn('Insert', 'primary', function () {
          WW.commands.exec('insert-table', { rows: rows, cols: cols, headerRow: headerRow });
          WW.events.changed('insert-table');
          close();
        })
      ]
    });
  }

  /* ============================================================ symbol == */
  var SYMBOLS = ('— – … ‘ ’ “ ” • ‣ ° ± × ÷ ½ ⅓ ¼ ¾ ‰ € £ ¥ © ® ™ § ¶ † ‡ ' +
    'α β γ δ ε π μ σ Ω Δ Σ ∞ ≠ ≤ ≥ ≈ ± √ ∫ → ← ↑ ↓ ↔ ⇔ ⇒ ∈ ∉ ⊂ ∪ ∅ ∀ ∃ ' +
    '✓ ✗ ★ ☆ ♥ ♦ ♣ ♠ ♪ ♫ ☺ ☀ ☂ ☃ ☎ ✂ ⌘ ⏎ ⇧ ␣ ① ② ③ ④ ⑤ ⑥ ⑦ ⑧ ⑨ ⑩ ' +
    'À Á Â Ä Å Æ Ç È É Ê Ë Ì Í Î Ï Ñ Ò Ó Ô Ö Ù Ú Û Ü ß Þ æ ö ü ñ').split(' ');

  function symbolDialog() {
    var insertAt = U.h('div', { class: 'hint', text: 'Click a character to insert it at the cursor.' });
    var grid = U.h('div', { class: 'palette', style: 'grid-template-columns:repeat(12,26px)' }, SYMBOLS.map(function (ch) {
      return U.h('button', {
        class: 'sw', type: 'button', text: ch,
        style: 'background:var(--bg-chrome);color:var(--text);font-size:14px;display:grid;place-items:center',
        onclick: function () {
          WW.commands.exec('insert-text', ch);
          WW.events.changed('insert-symbol');
          insertAt.textContent = 'Inserted “' + ch + '”.';
        }
      });
    }));
    show({ title: 'Special character', body: U.h('div', null, [grid, U.h('div', null, [insertAt])]), buttons: [btn('Close', '', close)] });
  }

  /* ======================================================== page setup == */
  function pageSetupDialog() {
    var p = WW.APP.page;
    var sizeSel = selectInput([['Letter', 'Letter 8.5×11″'], ['A4', 'A4 21×29.7cm'], ['Legal', 'Legal 8.5×14″'],
      ['A3', 'A3 29.7×42cm'], ['A5', 'A5 14.8×21cm'], ['Tabloid', 'Tabloid 11×17″'], ['B5', 'B5 17.6×25cm']], p.size);
    var orientSel = selectInput([['portrait', 'Portrait'], ['landscape', 'Landscape']], p.orientation);
    function numIn(v) { return U.h('input', { type: 'number', value: String(v), min: '0.25', max: '4', step: '0.25' }); }
    var mt = numIn(p.margins.top), mb = numIn(p.margins.bottom),
      ml = numIn(p.margins.left), mr = numIn(p.margins.right),
      mh = numIn(p.margins.header), mf = numIn(p.margins.footer);
    var body = U.h('div', null, [
      field('Paper size', sizeSel),
      field('Orientation', orientSel),
      U.h('h3', { text: 'Margins (inches)' }),
      U.h('div', { class: 'grid2' }, [
        field('Top', mt), field('Bottom', mb),
        field('Left', ml), field('Right', mr),
        field('Header', mh), field('Footer', mf)
      ])
    ]);
    show({
      title: 'Page setup', body: body,
      buttons: [
        btn('Cancel', '', close),
        btn('Apply', 'primary', function () {
          applyPageSetup({
            size: sizeSel.value, orientation: orientSel.value,
            margins: {
              top: parseFloat(mt.value) || 1, bottom: parseFloat(mb.value) || 1,
              left: parseFloat(ml.value) || 1, right: parseFloat(mr.value) || 1,
              header: parseFloat(mh.value) || 0.5, footer: parseFloat(mf.value) || 0.5
            }
          });
          close();
        })
      ]
    });
  }

  function applyPageSetup(p) {
    WW.APP.page = p;
    WW.pages.applyPage();
    U.$('#page-size').value = p.size;
    WW.events.changed('page-setup');
  }

  /* ========================================================== paragraph == */
  function paragraphDialog() {
    var b = WW.commands.anchorBlock();
    var cs = b ? window.getComputedStyle(b) : null;
    var align = (b && b.style.textAlign) || (cs && cs.textAlign) || 'left';
    var alignSel = selectInput([['left', 'Left'], ['center', 'Center'], ['right', 'Right'], ['justify', 'Justify']], align);
    function ptVal(prop) {
      if (!b || !b.style[prop]) return '';
      var m = /^([\d.]+)pt$/.exec(b.style[prop]);
      return m ? m[1] : '';
    }
    var before = textInput(ptVal('marginTop'), 'number');
    var after = textInput(ptVal('marginBottom'), 'number');
    var left = textInput(ptVal('marginLeft'), 'number');
    var right = textInput(ptVal('marginRight'), 'number');
    var first = textInput(b && /^-?[\d.]+pt$/.test(b.style.textIndent || '') ? parseFloat(b.style.textIndent) : '', 'number');
    var lineSel = selectInput([['', 'Default'], ['1', 'Single'], ['1.15', '1.15'], ['1.5', '1.5'], ['2', 'Double']],
      (b && b.style.lineHeight) || '');
    var body = U.h('div', null, [
      field('Alignment', alignSel),
      field('Line spacing', lineSel),
      U.h('h3', { text: 'Spacing (points)' }),
      U.h('div', { class: 'grid2' }, [field('Before', before), field('After', after)]),
      U.h('h3', { text: 'Indents (points)' }),
      U.h('div', { class: 'grid2' }, [
        field('Left', left), field('Right', right), field('First line', first)
      ])
    ]);
    show({
      title: 'Paragraph options', body: body,
      buttons: [
        btn('Cancel', '', close),
        btn('Apply', 'primary', function () {
          function num(el) { return el.value === '' ? null : parseFloat(el.value); }
          WW.commands.exec('setParagraphSpacing', {
            before: num(before), after: num(after),
            left: num(left), right: num(right), firstLine: num(first),
            align: alignSel.value
          });
          if (lineSel.value !== null) WW.commands.exec('setLineSpacing', lineSel.value);
          WW.events.changed('paragraph');
          close();
        })
      ]
    });
  }

  /* ========================================================= word count == */
  function wordCountDialog() {
    var text = (U.$('#editor').textContent || '').replace(/\u200b|\ufeff/g, '');
    var words = (text.match(/[^\s]+/g) || []).length;
    var chars = text.replace(/\s/g, '').length;
    var paras = U.$$('#editor > p, #editor > h1, #editor > h2, #editor > h3, #editor > h4, #editor > h5, #editor > h6, #editor > li').length;
    var heads = U.$$('#editor h1, #editor h2, #editor h3, #editor h4, #editor h5, #editor h6').length;
    var imgs = U.$$('#editor img').length;
    var tables = U.$$('#editor table').length;
    var read = Math.max(1, Math.ceil(words / 200));
    var rows = [
      ['Words', words], ['Characters (no spaces)', chars], ['Characters (with spaces)', text.length],
      ['Paragraphs', paras], ['Headings', heads], ['Images', imgs], ['Tables', tables],
      ['Pages', WW.pages.totalPages()], ['Reading time', read + ' min']
    ];
    var table = U.h('table', { class: 'kbd-table' }, rows.map(function (r) {
      return U.h('tr', null, [U.h('td', { text: r[0] }), U.h('td', { text: String(r[1]) })]);
    }));
    show({ title: 'Word count', body: table, buttons: [btn('Close', '', close)] });
  }

  /* ======================================================== preferences == */
  function preferencesDialog() {
    var st = WW.storage.getPrefs();
    var fontSel = selectInput(WW.APP.FONTS.map(function (f) { return [f, f]; }), st.defaultFont);
    var sizeSel = selectInput(WW.APP.FONT_SIZES.map(function (s) { return [String(s), s + ' pt']; }), st.defaultSize);
    var spell = U.h('input', { type: 'checkbox' }); spell.checked = !!st.spellcheck;
    var auto = U.h('input', { type: 'checkbox' }); auto.checked = !!st.autosave;
    var themeSel = selectInput([['light', 'Light'], ['dark', 'Dark']], st.theme);
    var body = U.h('div', null, [
      U.h('h3', { text: 'New documents' }),
      field('Default font', fontSel),
      field('Default size', sizeSel),
      U.h('h3', { text: 'Editing' }),
      U.h('div', { class: 'field' }, [U.h('label', { text: 'Spell check' }), spell]),
      U.h('div', { class: 'field' }, [U.h('label', { text: 'Autosave to this browser' }), auto]),
      U.h('div', { class: 'field' }, [U.h('label', { text: 'Theme' }), themeSel]),
      U.h('h3', { text: 'Storage' }),
      U.h('div', { class: 'field' }, [
        U.h('label', { text: 'Autosave copy' }),
        U.h('div', null, [
          btn('Clear autosave', 'danger', function () {
            WW.storage.clearAutosave();
            WW.events.toast('Autosave copy deleted', 'ok');
          }),
          U.h('span', { class: 'hint', text: '  Everything stays on this machine.' })
        ])
      ])
    ]);
    show({
      title: 'Preferences', body: body,
      buttons: [
        btn('Cancel', '', close),
        btn('Save', 'primary', function () {
          WW.storage.savePrefs({
            defaultFont: fontSel.value, defaultSize: parseFloat(sizeSel.value),
            spellcheck: spell.checked, autosave: auto.checked, theme: themeSel.value
          });
          WW.app.applyPrefs();
          WW.events.toast('Preferences saved', 'ok');
          close();
        })
      ]
    });
  }

  /* ============================================================ samples == */
  var SAMPLES = [
    {
      name: 'Project report',
      desc: 'Title page, headings, lists and a table of contents.',
      thumb: '<div style="font-size:1.6em;color:#2b579a">Project Report</div><div style="margin-top:.6em">Introduction<br>Approach<br>Results<br>Conclusion</div>',
      html: '<p class="ww-title">Project Report</p><p class="ww-subtitle">A sample document created with OpenWord</p>' +
        '<h1>Introduction</h1><p>This sample shows the building blocks of a report: headings, paragraphs, lists and tables. Everything here can be edited, saved as .docx or .odt, or printed to PDF.</p>' +
        '<h1>Approach</h1><ul><li>Gather requirements</li><li>Draft an outline</li><li>Iterate with stakeholders</li></ul>' +
        '<h1>Results</h1><p>The plan produced measurable improvements:</p>' +
        '<table style="width:100%"><thead><tr><th>Metric</th><th>Before</th><th>After</th></tr></thead>' +
        '<tbody><tr><td>Turnaround</td><td>5 days</td><td>2 days</td></tr><tr><td>Errors</td><td>12%</td><td>3%</td></tr></tbody></table>' +
        '<h1>Conclusion</h1><p>Use <em>Insert → Table of contents</em> to build a contents list from the headings above.</p>'
    },
    {
      name: 'Meeting minutes',
      desc: 'Structured notes with attendees and action items.',
      thumb: '<div style="font-size:1.4em;color:#2b579a">Minutes</div><div style="margin-top:.6em">Attendees:<br>• Ada<br>• Grace<br>• Alan</div>',
      html: '<h1>Meeting minutes</h1><p><b>Date:</b> <span class="ww-field" data-ww-field="date" contenteditable="false">today</span><br><b>Location:</b> Conference room B</p>' +
        '<h2>Attendees</h2><ul><li>Ada Lovelace</li><li>Grace Hopper</li><li>Alan Turing</li></ul>' +
        '<h2>Discussion</h2><p>The committee reviewed the draft specification and agreed on the naming scheme.</p>' +
        '<h2>Action items</h2><ol><li>Circulate the revised draft</li><li>Collect comments by Friday</li><li>Schedule the follow-up</li></ol>'
    },
    {
      name: 'Letter',
      desc: 'A classic business letter layout.',
      thumb: '<div style="text-align:right">12 Example Street</div><div style="margin-top:.6em">Dear Reader,<br><br>Just a quick note…</div>',
      html: '<p style="text-align:right">12 Example Street<br>Springfield<br><span class="ww-field" data-ww-field="date" contenteditable="false">today</span></p>' +
        '<p>Dear Reader,</p><p>Just a quick note to show how a plain business letter looks in OpenWord. The date above is a live field — it updates whenever the document is opened or saved.</p>' +
        '<p>Sincerely,</p><p>A. Writer</p>'
    }
  ];

  function samplesDialog() {
    var grid = U.h('div', { class: 'card-grid' }, SAMPLES.map(function (s) {
      return U.h('button', {
        class: 'card', type: 'button',
        onclick: function () {
          close();
          WW.app.loadDoc({ html: s.html, title: s.name, fresh: true });
        }
      }, [
        U.h('div', { class: 'thumb', html: s.thumb }),
        U.h('h4', { text: s.name }),
        U.h('p', { text: s.desc })
      ]);
    }));
    show({ title: 'Sample documents', body: grid, wide: true, buttons: [btn('Cancel', '', close)] });
  }

  /* ============================================================== gdoc == */
  function gdocDialog() {
    alert('Save for Google Docs',
      '<p>Your .docx has been downloaded and is ready for Google Docs.</p><h3>Upload steps</h3>' +
      '<ol><li>Open <b>drive.google.com</b> in your browser.</li>' +
      '<li>Click <b>New → File upload</b> and pick the downloaded .docx.</li>' +
      '<li>Once uploaded, right-click it → <b>Open with → Google Docs</b>.</li>' +
      '<li>In Google Docs use <b>File → Download → Microsoft Word (.docx)</b> to bring it back into OpenWord.</li></ol>' +
      '<p class="hint">OpenWord never uploads anything itself — the round trip is entirely file based.</p>');
  }

  /* ========================================================== shortcuts == */
  function shortcutsDialog() {
    var rows = (WW.shortcuts.LIST || []);
    var table = U.h('table', { class: 'kbd-table' }, rows.map(function (r) {
      return U.h('tr', null, [
        U.h('td', { html: r[0].split('+').map(function (k) { return '<kbd>' + U.escapeHtml(k) + '</kbd>'; }).join('+') }),
        U.h('td', { text: r[1] })
      ]);
    }));
    show({ title: 'Keyboard shortcuts', body: table, wide: true, buttons: [btn('Close', '', close)] });
  }

  function aboutDialog() {
    alert('About OpenWord',
      '<p><b>OpenWord</b> — a standalone, in-browser word processor.</p>' +
      '<p>DOCX and OpenDocument round-trip, HTML import/export, autosave and print — all running locally in your browser. No server, no accounts, no network traffic.</p>' +
      '<p class="hint">Built with JSZip and DOMPurify, both bundled locally.</p>');
  }

  /* ======================================================== import html == */
  function importHtmlDialog() {
    var ta = U.h('textarea', { placeholder: '<p>Paste HTML here…</p>' });
    var body = U.h('div', null, [
      U.h('p', { class: 'hint', text: 'The HTML is sanitised and normalised before it is inserted at the cursor.' }),
      ta
    ]);
    show({
      title: 'Import HTML', body: body, wide: true,
      buttons: [
        btn('Cancel', '', close),
        btn('Insert', 'primary', function () {
          var frag = WW.htmlio.fragmentFromHtml(ta.value);
          WW.commands.exec('insert-html', frag.innerHTML);
          WW.events.changed('import-html');
          close();
        })
      ]
    });
  }

  /* ============================================================ restore == */
  function restoreDialog(autosave) {
    return new Promise(function (resolve) {
      var when = new Date(autosave.ts || Date.now()).toLocaleString();
      show({
        title: 'Unsaved work found',
        body: U.h('div', null, [
          U.h('p', { text: 'OpenWord found an autosaved copy of “' + (autosave.title || 'Untitled document') + '” from ' + when + '.' }),
          U.h('p', { class: 'hint', text: 'Restore it, or start with a blank document?' })
        ]),
        buttons: [
          btn('Start fresh', 'danger', function () { resolve(false); close(); }),
          btn('Restore', 'primary', function () { resolve(true); close(); })
        ],
        onClose: function () { resolve(false); }
      });
    });
  }

  /* ------------------------------------------------------------- misc -- */
  function miniTabs(tabs) {
    var bar = U.h('div', { class: 'tabs-mini' });
    var host = U.h('div');
    tabs.forEach(function (t, i) {
      var b = U.h('button', { type: 'button', text: t[0], class: i === 0 ? 'active' : '' });
      b.addEventListener('click', function () {
        U.$$('button', bar).forEach(function (x) { x.classList.remove('active'); });
        b.classList.add('active');
        host.innerHTML = '';
        host.appendChild(t[1]);
      });
      bar.appendChild(b);
    });
    host.appendChild(tabs[0][1]);
    return U.h('div', null, [bar, host]);
  }

  /* ------------------------------------------------------------- init -- */
  function init() {
    backdrop = U.$('#modal-backdrop');
    modal = U.$('#modal');
    titleEl = U.$('#modal-title');
    bodyEl = U.$('#modal-body');
    footEl = U.$('#modal-foot');
    U.$('#modal-close').addEventListener('click', close);
    backdrop.addEventListener('mousedown', function (e) { if (e.target === backdrop) close(); });
  }

  function open(name) {
    var map = {
      'color': function () { colorDialog(false); },
      'highlight': function () { colorDialog(true); },
      'image': imageDialog, 'link': linkDialog, 'table': tableDialog,
      'symbol': symbolDialog, 'page-setup': pageSetupDialog, 'paragraph': paragraphDialog,
      'word-count': wordCountDialog, 'preferences': preferencesDialog, 'samples': samplesDialog,
      'gdoc': gdocDialog, 'shortcuts': shortcutsDialog, 'about': aboutDialog,
      'import-html': importHtmlDialog
    };
    if (map[name]) map[name]();
  }

  return {
    init: init, open: open, show: show, close: close, isOpen: isOpen,
    confirm: confirm, alert: alert, restoreDialog: restoreDialog,
    applyPageSetup: applyPageSetup, SAMPLES: SAMPLES
  };
})(WW.util);
