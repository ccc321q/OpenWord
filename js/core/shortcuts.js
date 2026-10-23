/* ==========================================================================
   OpenWord — keyboard shortcuts.
   Exposes WW.shortcuts. The LIST doubles as the data source for the
   Help → Keyboard shortcuts dialog, so the map and the docs never drift.
   ========================================================================== */
WW.shortcuts = (function (U) {
  'use strict';

  var LIST = [
    ['Ctrl+B', 'Bold'], ['Ctrl+I', 'Italic'], ['Ctrl+U', 'Underline'],
    ['Ctrl+Shift+X', 'Strikethrough'], ['Ctrl+Space', 'Clear formatting'],
    ['Ctrl+Z', 'Undo'], ['Ctrl+Y', 'Redo'],
    ['Ctrl+X', 'Cut'], ['Ctrl+C', 'Copy'], ['Ctrl+V', 'Paste'],
    ['Ctrl+Shift+V', 'Paste plain text'], ['Ctrl+A', 'Select all'],
    ['Ctrl+F', 'Find'], ['Ctrl+H', 'Find & Replace'], ['Esc', 'Close find / dialog / menu'],
    ['Ctrl+K', 'Insert link'], ['Ctrl+Enter', 'Page break'],
    ['Tab', 'Indent / next table cell'], ['Shift+Tab', 'Outdent / previous cell'],
    ['Ctrl+Shift+L', 'Align left'], ['Ctrl+Shift+E', 'Center'],
    ['Ctrl+Shift+R', 'Align right'], ['Ctrl+J', 'Justify'],
    ['Ctrl+Shift+7', 'Bulleted list'], ['Ctrl+Shift+8', 'Numbered list'],
    ['Ctrl+=', 'Subscript'], ['Ctrl+Shift++', 'Superscript'],
    ['Ctrl+Alt+C', 'Copy formatting (painter)'], ['Ctrl+Alt+V', 'Paste formatting'],
    ['Ctrl+S', 'Save as .docx'], ['Ctrl+O', 'Open file'], ['Ctrl+P', 'Print'],
    ['Ctrl+Alt+N', 'New document'], ['F1', 'Keyboard shortcuts'],
    ['Alt', 'Focus the menu bar']
  ];

  function combo(e) {
    var k = e.key;
    if (k === 'Control' || k === 'Shift' || k === 'Alt' || k === 'Meta') return null;
    if (k.length === 1) k = k.toLowerCase();
    return (e.ctrlKey || e.metaKey ? 'ctrl+' : '') +
      (e.altKey ? 'alt+' : '') +
      (e.shiftKey ? 'shift+' : '') + k;
  }

  function inTextField(el) {
    return el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.tagName === 'SELECT' ||
      el.isContentEditable);
  }

  /* ------------------------------------------------------- tab handling -- */
  function handleTab(e, shift) {
    var sel = window.getSelection();
    var node = sel && sel.anchorNode;
    var editable = node && (node.nodeType === 1 ? node : node.parentElement);
    editable = editable && editable.closest ? editable.closest('[contenteditable="true"]') : null;
    if (!editable) return false;

    var cell = node && (node.nodeType === 1 ? node : node.parentElement).closest &&
      (node.nodeType === 1 ? node : node.parentElement).closest('td,th');
    if (cell) { moveCell(cell, shift ? -1 : 1); e.preventDefault(); return true; }

    var li = WW.commands.listInfoOf(WW.commands.anchorBlock());
    if (li) {
      e.preventDefault(); // never let Tab leave the editor inside a list
      if (shift) { if (WW.commands.changeListIndent(-1)) WW.events.changed('outdent'); }
      else if (li.li.previousElementSibling) {
        if (WW.commands.changeListIndent(1)) WW.events.changed('indent');
      }
      return true;
    }
    // plain tab character
    e.preventDefault();
    WW.commands.exec('insert-text', '\t');
    WW.events.changed('tab');
    return true;
  }

  function moveCell(cell, dir) {
    var table = cell.closest('table');
    var row = cell.parentElement;
    var cells = Array.prototype.slice.call(table.querySelectorAll('td,th'));
    var i = cells.indexOf(cell);
    var next = cells[i + dir];
    if (!next && dir > 0) {
      // at the end: add a new row like Word does
      var tr = document.createElement('tr');
      for (var c = 0; c < row.cells.length; c++) {
        var td = document.createElement(row.cells[0].tagName.toLowerCase());
        td.innerHTML = '<p><br></p>';
        tr.appendChild(td);
      }
      row.parentElement.appendChild(tr);
      next = tr.cells[0];
    }
    if (!next) return;
    var range = document.createRange();
    range.selectNodeContents(next);
    range.collapse(dir > 0);
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    WW.pages.relayout();
  }

  /* -------------------------------------------------------------- map -- */
  function map() {
    var C = WW.commands, A = WW.app;
    var m = {
      'ctrl+b': { run: function () { C.exec('bold'); WW.events.changed('bold'); } },
      'ctrl+i': { run: function () { C.exec('italic'); WW.events.changed('italic'); } },
      'ctrl+u': { run: function () { C.exec('underline'); WW.events.changed('underline'); } },
      'ctrl+shift+x': { run: function () { C.exec('strikeThrough'); WW.events.changed('strikeThrough'); } },
      'ctrl+space': { run: function () { C.exec('clear-formatting'); WW.events.changed('clear-formatting'); } },
      'ctrl+z': { run: function () { WW.history.undo(); } },
      'ctrl+y': { run: function () { WW.history.redo(); } },
      'ctrl+shift+z': { run: function () { WW.history.redo(); } },
      'ctrl+f': { run: function () { A.openFind(false); } },
      'ctrl+h': { run: function () { A.openFind(true); } },
      'ctrl+k': { run: function () { WW.dialogs.open('link'); } },
      'ctrl+enter': { run: function () { C.exec('insert-page-break'); WW.events.changed('insert-page-break'); } },
      'ctrl+shift+v': { run: function () { WW.menus.clipboardPaste(); } },
      'ctrl+shift+l': { run: function () { C.exec('justifyLeft'); WW.events.changed('align'); } },
      'ctrl+shift+e': { run: function () { C.exec('justifyCenter'); WW.events.changed('align'); } },
      'ctrl+shift+r': { run: function () { C.exec('justifyRight'); WW.events.changed('align'); } },
      'ctrl+j': { run: function () { C.exec('justifyFull'); WW.events.changed('align'); } },
      'ctrl+shift+7': { run: function () { C.exec('insertUnorderedList'); WW.events.changed('list'); } },
      'ctrl+shift+8': { run: function () { C.exec('insertOrderedList'); WW.events.changed('list'); } },
      'ctrl+=': { run: function () { C.exec('subscript'); WW.events.changed('subscript'); } },
      'ctrl+shift++': { run: function () { C.exec('superscript'); WW.events.changed('superscript'); } },
      'ctrl+alt+c': { run: function () { C.capturePainter(); WW.events.toast('Formatting copied', 'ok'); } },
      'ctrl+alt+v': { run: function () { C.applyPainter(); } },
      'ctrl+s': { run: function () { A.runCmd('save-docx'); } },
      'ctrl+o': { run: function () { A.runCmd('open'); } },
      'ctrl+p': { run: function () { A.runCmd('print'); } },
      'ctrl+alt+n': { run: function () { A.runCmd('new'); } },
      'f1': { run: function () { WW.dialogs.open('shortcuts'); } },
      'escape': { run: function () { A.escape(); } }
    };
    return m;
  }

  /* -------------------------------------------------------------- init -- */
  function init() {
    var M = map();
    document.addEventListener('keydown', function (e) {
      // Tab is handled before the generic map so the browser never moves focus
      if (e.key === 'Tab' && !e.ctrlKey && !e.altKey && !e.metaKey) {
        handleTab(e, e.shiftKey);
        return;
      }
      var c = combo(e);
      if (!c) {
        if (e.key === 'Alt' && !e.ctrlKey && !e.shiftKey && !e.metaKey) {
          setTimeout(function () {
            var first = document.querySelector('.menu-label');
            if (first && !WW.menus.isOpen()) first.focus();
          }, 0);
        }
        return;
      }
      var entry = M[c];
      if (!entry) return;
      // inside dialogs, only Esc is intercepted
      if (WW.dialogs.isOpen() && c !== 'escape') return;
      e.preventDefault();
      entry.run();
    }, true);
  }

  return { init: init, LIST: LIST };
})(WW.util);
