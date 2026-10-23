/* ==========================================================================
   OpenWord — find & replace
   Uses the CSS Custom Highlight API when available (zero DOM churn) and always
   keeps a real selection on the active match so that replace/paste behave.
   ========================================================================== */
WW.find = (function (U, Sel) {
  'use strict';

  var matches = [];
  var index = -1;
  var opts = { caseSensitive: false, wholeWord: false };
  var supportsHighlight = (typeof window.Highlight === 'function') && window.CSS && CSS.highlights;
  var listeners = [];

  function editor() { return document.getElementById('editor'); }

  function textNodes(root) {
    var out = [];
    var walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: function (node) {
        if (!node.nodeValue || !node.nodeValue.length) return NodeFilter.FILTER_REJECT;
        var p = node.parentElement;
        if (!p) return NodeFilter.FILTER_REJECT;
        if (p.closest('.ww-nosrc, style, script')) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    var n;
    while ((n = walker.nextNode())) out.push(n);
    return out;
  }

  function buildMatcher(term) {
    var flags = 'g' + (opts.caseSensitive ? '' : 'i');
    var pattern;
    if (opts.wholeWord) {
      pattern = '\\b' + term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b';
    } else {
      return { literal: !opts.caseSensitive ? null : term, re: new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags) };
    }
    return { re: new RegExp(pattern, flags) };
  }

  function search(term, options) {
    if (options) {
      opts.caseSensitive = !!options.caseSensitive;
      opts.wholeWord = !!options.wholeWord;
    }
    clearHighlights();
    matches = [];
    index = -1;
    if (!term) { emit(); return matches; }
    var root = editor();
    if (!root) return matches;
    var nodes = textNodes(root);
    var matcher = buildMatcher(term);
    nodes.forEach(function (node) {
      var text = node.nodeValue;
      matcher.re.lastIndex = 0;
      var m;
      while ((m = matcher.re.exec(text))) {
        if (m[0].length === 0) { matcher.re.lastIndex++; continue; }
        var range = document.createRange();
        try {
          range.setStart(node, m.index);
          range.setEnd(node, m.index + m[0].length);
        } catch (e) { continue; }
        matches.push(range);
        if (matches.length > 5000) return;
      }
    });
    if (supportsHighlight) {
      var hl = new Highlight();
      matches.forEach(function (r) { hl.add(r); });
      CSS.highlights.set('ww-find', hl);
    }
    if (matches.length) focusMatch(0);
    emit();
    return matches;
  }

  function focusMatch(i) {
    if (!matches.length) return;
    index = ((i % matches.length) + matches.length) % matches.length;
    if (supportsHighlight) {
      var active = new Highlight();
      active.add(matches[index]);
      CSS.highlights.set('ww-find-active', active);
    }
    var range = matches[index];
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    Sel.scrollCaretIntoView();
    emit();
  }

  function next() { if (matches.length) focusMatch(index + 1); }
  function prev() { if (matches.length) focusMatch(index - 1); }

  function currentTermGetter(term) { return term; }

  function replaceOne(term, replacement) {
    if (index < 0 || !matches.length) return false;
    var range = matches[index];
    var sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    editor().focus();
    var ok = document.execCommand('insertText', false, replacement);
    if (!ok) {
      // Fallback: direct DOM surgery.
      try {
        range.deleteContents();
        range.insertNode(document.createTextNode(replacement));
      } catch (e) { return false; }
    }
    if (WW.history) WW.history.push('replace');
    var newIndex = Math.min(index, Math.max(0, matches.length - 2));
    search(term);
    if (matches.length) focusMatch(newIndex);
    return true;
  }

  function replaceAll(term, replacement) {
    if (!matches.length) return 0;
    var count = matches.length;
    // Walk backwards so earlier ranges stay valid while we edit.
    for (var i = matches.length - 1; i >= 0; i--) {
      var range = matches[i];
      try {
        range.deleteContents();
        if (replacement) range.insertNode(document.createTextNode(replacement));
      } catch (e) { /* stale range */ }
    }
    if (WW.history) WW.history.push('replace all');
    search(term);
    if (matches.length) focusMatch(0);
    return count;
  }

  function clearHighlights() {
    if (supportsHighlight) {
      CSS.highlights.delete('ww-find');
      CSS.highlights.delete('ww-find-active');
    }
  }

  function clear() {
    clearHighlights();
    matches = [];
    index = -1;
    emit();
  }

  function stats() {
    return { total: matches.length, index: index };
  }

  function onChange(fn) { listeners.push(fn); }
  function emit() { listeners.forEach(function (fn) { try { fn(stats()); } catch (e) { } }); }

  return {
    search: search, next: next, prev: prev, replaceOne: replaceOne, replaceAll: replaceAll,
    clear: clear, stats: stats, onChange: onChange, clearHighlights: clearHighlights
  };
})(WW.util, WW.selection);
