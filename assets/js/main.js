/* Application-Logs-Viewer: theme toggle, copy buttons, command builder, contents highlight */
(function () {
  var $ = function (id) { return document.getElementById(id); };

  /* ---------- theme ---------- */
  var sun = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>';
  var moon = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>';
  var root = document.documentElement;
  function effective() {
    var t = root.getAttribute('data-theme');
    if (t) return t;
    return window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }
  function paintToggle() { $('theme').innerHTML = effective() === 'dark' ? sun : moon; }
  try { var saved = localStorage.getItem('eventreport-theme'); if (saved === 'dark' || saved === 'light') root.setAttribute('data-theme', saved); } catch (e) {}
  paintToggle();
  $('theme').addEventListener('click', function () {
    var next = effective() === 'dark' ? 'light' : 'dark';
    root.setAttribute('data-theme', next);
    try { localStorage.setItem('eventreport-theme', next); } catch (e) {}
    paintToggle();
  });

  /* ---------- copy buttons ---------- */
  function fallbackCopy(text, pre) {
    var ok = false;
    try {
      var ta = document.createElement('textarea');
      ta.value = text; ta.setAttribute('readonly', ''); ta.style.cssText = 'position:fixed;top:0;left:0;opacity:0';
      document.body.appendChild(ta); ta.select(); ok = document.execCommand('copy'); document.body.removeChild(ta);
    } catch (e) { ok = false; }
    if (!ok && pre && window.getSelection) {
      var r = document.createRange(); r.selectNodeContents(pre);
      var s = getSelection(); s.removeAllRanges(); s.addRange(r);
    }
    return ok;
  }
  function flash(btn, label, good) {
    btn.textContent = label; btn.classList.toggle('done', !!good);
    setTimeout(function () { btn.textContent = 'Copy'; btn.classList.remove('done'); }, 1800);
  }
  Array.prototype.forEach.call(document.querySelectorAll('.code'), function (box) {
    var pre = box.querySelector('pre');
    var btn = document.createElement('button');
    btn.type = 'button'; btn.className = 'copy'; btn.textContent = 'Copy';
    btn.setAttribute('aria-label', 'Copy command');
    btn.addEventListener('click', function () {
      var text = pre.textContent.trim();
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text).then(function () { flash(btn, 'Copied', true); }, function () {
          flash(btn, fallbackCopy(text, pre) ? 'Copied' : 'Press Ctrl+C', true);
        });
      } else {
        flash(btn, fallbackCopy(text, pre) ? 'Copied' : 'Press Ctrl+C', true);
      }
    });
    box.appendChild(btn);
  });

  /* ---------- command builder ---------- */
  var fStart = $('f-start'), fEnd = $('f-end'), fFolder = $('f-folder'), fCrit = $('f-crit'), fOpen = $('f-open');
  function quote(p) { return /[\s"]/.test(p) ? '"' + p.replace(/"/g, '`"') + '"' : p; }
  function build() {
    var parts = ['.\\Get-AppEventReport.ps1'];
    if (fStart.value) parts.push('-StartDate ' + fStart.value);
    if (fEnd.value) parts.push('-EndDate ' + fEnd.value);
    var folder = fFolder.value.trim();
    if (folder) parts.push('-OutputFolder ' + quote(folder));
    if (fCrit.checked) parts.push('-IncludeCritical');
    if (fOpen.checked) parts.push('-OpenReport');
    $('built').textContent = parts.join(' ');
    var note = $('built-note'), msg = '', bad = false;
    if (fStart.value && fEnd.value && fEnd.value < fStart.value) { msg = 'The end date is before the start date. The script will stop with an error.'; bad = true; }
    else if (!fStart.value && !fEnd.value) msg = 'No dates set. PowerShell will ask for both when the script starts.';
    else if (!fStart.value) msg = 'No start date set. PowerShell will ask for it, with 7 days ago as the default.';
    else if (!fEnd.value) msg = 'No end date set. PowerShell will ask for it, with today as the default.';
    else msg = 'Both dates are set, so the script runs without any prompt.';
    note.textContent = msg; note.classList.toggle('bad', bad);
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function iso(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  Array.prototype.forEach.call(document.querySelectorAll('[data-preset]'), function (b) {
    b.addEventListener('click', function () {
      var today = new Date(), start = new Date(today), k = b.getAttribute('data-preset');
      if (k === '7') start.setDate(start.getDate() - 7);
      else if (k === '30') start.setDate(start.getDate() - 30);
      else if (k === 'month') start.setDate(1);
      fStart.value = iso(start); fEnd.value = iso(today); build();
    });
  });
  [fStart, fEnd, fFolder, fCrit, fOpen].forEach(function (el) { el.addEventListener('input', build); el.addEventListener('change', build); });
  build();

  /* ---------- contents highlight ---------- */
  var links = Array.prototype.slice.call(document.querySelectorAll('#toc a'));
  var byId = {}; links.forEach(function (a) { byId[a.getAttribute('href').slice(1)] = a; });
  function setActive(id) {
    links.forEach(function (a) { a.classList.toggle('on', a === byId[id]); if (a === byId[id]) a.setAttribute('aria-current', 'true'); else a.removeAttribute('aria-current'); });
    var a = byId[id], strip = $('toc');
    if (a && strip.scrollWidth > strip.clientWidth) { strip.scrollLeft = a.offsetLeft - (strip.clientWidth - a.offsetWidth) / 2; }
  }
  if ('IntersectionObserver' in window) {
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) { if (e.isIntersecting) setActive(e.target.id); });
    }, { rootMargin: '-25% 0px -65% 0px', threshold: 0 });
    Array.prototype.forEach.call(document.querySelectorAll('main section[id]'), function (s) { io.observe(s); });
  }
  setActive('overview');
})();
