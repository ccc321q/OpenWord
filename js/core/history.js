/* ==========================================================================
   OpenWord — undo / redo
   Snapshot based (innerHTML + serialised selection). Snapshots are cheap
   because documents are small, and they never leave the editor out of sync the
   way native execCommand undo does once we mutate the DOM ourselves.
   ========================================================================== */
WW.history = (function (U, Sel) {
  'use strict';

  var LIMIT = 150;
  var undoStack = [];
  var redoStack = [];
  var editor, header, footer;
  var last = { editor: '', header: '', footer: '', sel: null, label: '' };
  var suspended = false;
  var listeners = [];

  function init() {
    editor = document.getElementById('editor');
    header = document.getElementById('page-header');
    footer = document.getElementById('page-footer');
    reset('');
  }

  function snapshot() {
    return {
      editor: editor ? editor.innerHTML : '',
      header: header ? header.innerHTML : '',
      footer: footer ? footer.innerHTML : '',
      sel: Sel.capture()
    };
  }

  function sameContent(a, b) {
    return a === b.editor;
  }

  /** Record the *current* DOM as a new state; the previous state becomes undoable. */
  function push(label) {
    if (suspended) return false;
    if (!editor) init();
    var snap = snapshot();
    if (snap.editor === last.editor && snap.header === last.header && snap.footer === last.footer) {
      last.sel = snap.sel || last.sel;
      last.label = label || last.label;
      return false;
    }
    undoStack.push({
      editor: last.editor, header: last.header, footer: last.footer,
      sel: last.sel, label: last.label || 'typing'
    });
    if (undoStack.length > LIMIT) undoStack.shift();
    redoStack.length = 0;
    last = {
      editor: snap.editor, header: snap.header, footer: snap.footer,
      sel: snap.sel, label: label || 'typing'
    };
    emit();
    return true;
  }

  var debounced = U.debounce(function () { push('typing'); }, 550);
  function pushDebounced(label) { last.label = label || last.label; debounced(); }
  function flush() { debounced.flush(); }

  function applyState(entry) {
    suspended = true;
    editor.innerHTML = entry.editor;
    if (header) header.innerHTML = entry.header;
    if (footer) footer.innerHTML = entry.footer;
    suspended = false;
    last = {
      editor: entry.editor, header: entry.header, footer: entry.footer,
      sel: entry.sel, label: entry.label
    };
    Sel.restore(entry.sel);
    emit();
  }

  function undo() {
    flush();
    if (!undoStack.length) return false;
    var cur = snapshot();
    redoStack.push({ editor: cur.editor, header: cur.header, footer: cur.footer, sel: cur.sel, label: last.label });
    applyState(undoStack.pop());
    if (WW.events) WW.events.changed('undo');
    return true;
  }

  function redo() {
    if (!redoStack.length) return false;
    var cur = snapshot();
    undoStack.push({ editor: last.editor, header: last.header, footer: last.footer, sel: last.sel, label: last.label });
    applyState(redoStack.pop());
    if (WW.events) WW.events.changed('redo');
    return true;
  }

  function reset(label) {
    undoStack.length = 0;
    redoStack.length = 0;
    if (!editor) init();
    var snap = snapshot();
    last = { editor: snap.editor, header: snap.header, footer: snap.footer, sel: snap.sel, label: label || '' };
    emit();
  }

  function canUndo() { return undoStack.length > 0; }
  function canRedo() { return redoStack.length > 0; }
  function labels() { return { undo: undoStack.map(function (e) { return e.label; }), redo: redoStack.map(function (e) { return e.label; }) }; }

  function onChange(fn) { listeners.push(fn); }
  function emit() { listeners.forEach(function (fn) { try { fn(canUndo(), canRedo()); } catch (e) { } }); }

  function suspend(fn) {
    suspended = true;
    try { fn(); } finally {
      suspended = false;
      var snap = snapshot();
      last.editor = snap.editor; last.header = snap.header; last.footer = snap.footer;
    }
  }

  return {
    init: init, push: push, pushDebounced: pushDebounced, flush: flush,
    undo: undo, redo: redo, reset: reset, canUndo: canUndo, canRedo: canRedo,
    onChange: onChange, suspend: suspend, labels: labels,
    clear: function () { undoStack.length = 0; redoStack.length = 0; emit(); }
  };
})(WW.util, WW.selection);
