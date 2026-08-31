/* =============================================================================
   «Сигналы и протоколы» — интерактивный слой.
   Frontend-only. Никаких сетевых запросов, трекеров и передачи данных.
   Состояние чек-листов хранится ТОЛЬКО в localStorage браузера пользователя.
   ============================================================================= */
(function () {
  'use strict';

  var STORE_KEY = 'sp-protocol-state-v1';
  var reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* ---------- persistence (local only) ---------------------------------- */
  var state = { steps: {}, vendor: {}, gate: {} };

  function loadState() {
    try {
      var raw = window.localStorage.getItem(STORE_KEY);
      if (!raw) return;
      var parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') {
        state.steps  = parsed.steps  || {};
        state.vendor = parsed.vendor || {};
        state.gate   = parsed.gate   || {};
      }
    } catch (e) { /* приватный режим или отключённое хранилище — работаем без сохранения */ }
  }

  function saveState() {
    try { window.localStorage.setItem(STORE_KEY, JSON.stringify(state)); }
    catch (e) { /* тихо игнорируем: функциональность страницы от этого не зависит */ }
  }

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  /* ---------- 1. reading progress --------------------------------------- */
  function initProgress() {
    var bar = $('#progress-bar');
    if (!bar) return;
    var ticking = false;
    function update() {
      var h = document.documentElement;
      var max = h.scrollHeight - h.clientHeight;
      var pct = max > 0 ? Math.min(100, Math.max(0, (h.scrollTop || window.pageYOffset) / max * 100)) : 0;
      bar.style.width = pct.toFixed(1) + '%';
      bar.parentNode.setAttribute('aria-valuenow', Math.round(pct));
      ticking = false;
    }
    window.addEventListener('scroll', function () {
      if (!ticking) { ticking = true; window.requestAnimationFrame(update); }
    }, { passive: true });
    update();
  }

  /* ---------- 2. active section in table of contents --------------------- */
  function initTocObserver() {
    var links = $$('.toc a[href^="#"]');
    if (!links.length || !('IntersectionObserver' in window)) return;
    var byId = {};
    links.forEach(function (a) { byId[a.getAttribute('href').slice(1)] = a; });

    var visible = {};
    var obs = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        visible[en.target.id] = en.isIntersecting ? en.intersectionRatio : 0;
      });
      var bestId = null, best = 0;
      Object.keys(visible).forEach(function (id) {
        if (visible[id] > best) { best = visible[id]; bestId = id; }
      });
      links.forEach(function (a) { a.removeAttribute('aria-current'); });
      if (bestId && byId[bestId]) byId[bestId].setAttribute('aria-current', 'true');
    }, { rootMargin: '-72px 0px -55% 0px', threshold: [0, 0.05, 0.25, 0.6, 1] });

    Object.keys(byId).forEach(function (id) {
      var el = document.getElementById(id);
      if (el) obs.observe(el);
    });
  }

  /* ---------- 3. protocol steps highlight on scroll ---------------------- */
  function initStepObserver() {
    var steps = $$('.step');
    if (!steps.length || !('IntersectionObserver' in window) || reduceMotion) return;
    var obs = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) { en.target.classList.toggle('is-inview', en.isIntersecting); });
    }, { rootMargin: '-25% 0px -35% 0px', threshold: 0 });
    steps.forEach(function (s) { obs.observe(s); });
  }

  /* ---------- 4. mobile navigation --------------------------------------- */
  function initNavToggle() {
    var btn = $('#nav-toggle'), panel = $('#index-col');
    if (!btn || !panel) return;

    function isMobile() { return window.matchMedia('(max-width: 1080px)').matches; }
    function setOpen(open) {
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
      if (open) { panel.removeAttribute('hidden'); } else { panel.setAttribute('hidden', ''); }
      btn.querySelector('.nav-toggle-label').textContent = open ? 'Закрыть' : 'Разделы';
    }
    function sync() {
      if (isMobile()) { setOpen(btn.getAttribute('aria-expanded') === 'true'); }
      else { panel.removeAttribute('hidden'); btn.setAttribute('aria-expanded', 'false'); }
    }
    btn.addEventListener('click', function () { setOpen(btn.getAttribute('aria-expanded') !== 'true'); });
    panel.addEventListener('click', function (e) {
      if (e.target.closest('a') && isMobile()) setOpen(false);
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && isMobile() && btn.getAttribute('aria-expanded') === 'true') {
        setOpen(false); btn.focus();
      }
    });
    window.addEventListener('resize', sync);
    if (isMobile()) setOpen(false);
  }

  /* ---------- 5. generic disclosure (risks, flow nodes) ------------------ */
  function initDisclosures() {
    $$('[data-toggle]').forEach(function (btn) {
      var target = document.getElementById(btn.getAttribute('aria-controls'));
      if (!target) return;
      btn.addEventListener('click', function () {
        var open = btn.getAttribute('aria-expanded') === 'true';
        btn.setAttribute('aria-expanded', open ? 'false' : 'true');
        if (open) { target.setAttribute('hidden', ''); } else { target.removeAttribute('hidden'); }
      });
    });
  }

  /* ---------- 6. filters (risk matrix + data map) ------------------------ */
  function initFilterGroup(groupSel) {
    var group = $(groupSel);
    if (!group) return;
    var listSel = group.getAttribute('data-target');
    var items = $$(listSel + ' [data-filter]');
    var counter = $('[data-count]', group);
    var chips = $$('.chip', group);

    function apply(value) {
      var shown = 0;
      items.forEach(function (item) {
        var tags = (item.getAttribute('data-filter') || '').split(/\s+/);
        var match = value === 'all' || tags.indexOf(value) !== -1;
        item.classList.toggle('is-hidden', !match);
        if (match) shown++;
      });
      chips.forEach(function (c) { c.setAttribute('aria-pressed', c.getAttribute('data-value') === value ? 'true' : 'false'); });
      if (counter) counter.textContent = 'Показано: ' + shown + ' из ' + items.length;
    }

    chips.forEach(function (c) {
      c.addEventListener('click', function () { apply(c.getAttribute('data-value')); });
    });
    apply('all');
  }

  /* ---------- 7. glossary search ----------------------------------------- */
  function initGlossary() {
    var input = $('#glossary-search');
    var list = $('#glossary-list');
    if (!input || !list) return;
    var terms = $$('.term', list);
    var empty = $('#glossary-empty');
    var counter = $('#glossary-count');

    function normalize(s) { return (s || '').toLowerCase().replace(/ё/g, 'е').trim(); }

    function run() {
      var q = normalize(input.value);
      var shown = 0;
      terms.forEach(function (t) {
        var hay = normalize(t.getAttribute('data-search') + ' ' + t.textContent);
        var match = !q || hay.indexOf(q) !== -1;
        t.classList.toggle('is-hidden', !match);
        if (match) shown++;
        if (q && match && q.length > 2) { t.setAttribute('open', ''); }
        else if (!q) { t.removeAttribute('open'); }
      });
      if (empty) { empty.classList.toggle('is-hidden', shown !== 0); }
      if (counter) counter.textContent = shown + ' / ' + terms.length;
    }

    input.addEventListener('input', run);
    var clear = $('#glossary-clear');
    if (clear) clear.addEventListener('click', function () { input.value = ''; run(); input.focus(); });
    run();
  }

  /* ---------- 8. protocol checklist (localStorage) ----------------------- */
  function initSteps() {
    var steps = $$('.step');
    if (!steps.length) return;
    var meter = $('#protocol-meter > i');
    var label = $('#protocol-progress-label');

    function refresh() {
      var done = steps.filter(function (s) { return s.classList.contains('is-done'); }).length;
      if (meter) meter.style.width = (done / steps.length * 100) + '%';
      if (label) label.textContent = done + ' из ' + steps.length + ' этапов отмечено';
    }

    steps.forEach(function (step) {
      var id = step.getAttribute('data-step');
      var btn = $('[data-step-toggle]', step);
      var stateEl = $('[data-step-state]', step);
      if (!btn) return;

      function paint() {
        var done = !!state.steps[id];
        step.classList.toggle('is-done', done);
        btn.setAttribute('aria-pressed', done ? 'true' : 'false');
        btn.textContent = done ? 'Снять отметку' : 'Отметить выполненным';
        if (stateEl) stateEl.textContent = done ? 'СТАТУС: ВЫПОЛНЕНО (локально)' : 'СТАТУС: НЕ ОТМЕЧЕНО';
        refresh();
      }
      btn.addEventListener('click', function () {
        state.steps[id] = !state.steps[id];
        saveState(); paint();
      });
      paint();
    });

    var reset = $('#protocol-reset');
    if (reset) reset.addEventListener('click', function () {
      state.steps = {}; saveState();
      steps.forEach(function (step) {
        var btn = $('[data-step-toggle]', step);
        var stateEl = $('[data-step-state]', step);
        step.classList.remove('is-done');
        if (btn) { btn.setAttribute('aria-pressed', 'false'); btn.textContent = 'Отметить выполненным'; }
        if (stateEl) stateEl.textContent = 'СТАТУС: НЕ ОТМЕЧЕНО';
      });
      refresh();
    });
    refresh();
  }

  /* ---------- 9. vendor due diligence checklist -------------------------- */
  function initVendorList() {
    var boxes = $$('#vendor-list .vendor-item input[type="checkbox"]');
    if (!boxes.length) return;
    var meter = $('#vendor-meter > i');
    var label = $('#vendor-progress-label');

    function refresh() {
      var done = boxes.filter(function (b) { return b.checked; }).length;
      if (meter) meter.style.width = (done / boxes.length * 100) + '%';
      if (label) label.textContent = done + ' из ' + boxes.length + ' пунктов подтверждено документом';
    }

    boxes.forEach(function (box) {
      var id = box.id;
      box.checked = !!state.vendor[id];
      box.closest('.vendor-item').classList.toggle('is-checked', box.checked);
      box.addEventListener('change', function () {
        state.vendor[id] = box.checked;
        box.closest('.vendor-item').classList.toggle('is-checked', box.checked);
        saveState(); refresh();
      });
    });

    var reset = $('#vendor-reset');
    if (reset) reset.addEventListener('click', function () {
      state.vendor = {}; saveState();
      boxes.forEach(function (b) { b.checked = false; b.closest('.vendor-item').classList.remove('is-checked'); });
      refresh();
    });
    refresh();
  }

  /* ---------- 10. copy buttons ------------------------------------------- */
  function initCopy() {
    $$('[data-copy]').forEach(function (btn) {
      var out = btn.parentNode.querySelector('.copy-state');
      btn.addEventListener('click', function () {
        var src = document.getElementById(btn.getAttribute('data-copy'));
        if (!src) return;
        var text = src.textContent.trim();
        var done = function (ok) {
          if (!out) return;
          out.textContent = ok ? '✓ скопировано' : 'выделите текст вручную';
          window.setTimeout(function () { out.textContent = ''; }, 2600);
        };
        if (navigator.clipboard && window.isSecureContext) {
          navigator.clipboard.writeText(text).then(function () { done(true); }, function () { fallback(text, done); });
        } else { fallback(text, done); }
      });
    });

    function fallback(text, done) {
      try {
        var ta = document.createElement('textarea');
        ta.value = text;
        ta.setAttribute('readonly', '');
        ta.style.position = 'fixed'; ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.select();
        var ok = document.execCommand('copy');
        document.body.removeChild(ta);
        done(ok);
      } catch (e) { done(false); }
    }
  }

  /* ---------- 11. quizzes ------------------------------------------------ */
  function initQuiz() {
    $$('.quiz').forEach(function (quiz) {
      var explain = $('.quiz-explain', quiz);
      var opts = $$('.quiz-opt', quiz);
      opts.forEach(function (opt) {
        opt.addEventListener('click', function () {
          var correct = opt.getAttribute('data-correct') === 'true';
          opts.forEach(function (o) {
            o.classList.remove('is-correct', 'is-wrong');
            o.setAttribute('aria-pressed', 'false');
            var m = o.querySelector('.mark'); if (m) m.textContent = '□';
          });
          opt.setAttribute('aria-pressed', 'true');
          opt.classList.add(correct ? 'is-correct' : 'is-wrong');
          var mark = opt.querySelector('.mark');
          if (mark) mark.textContent = correct ? '✓' : '×';
          if (!correct) {
            opts.forEach(function (o) {
              if (o.getAttribute('data-correct') === 'true') {
                o.classList.add('is-correct');
                var m2 = o.querySelector('.mark'); if (m2) m2.textContent = '✓';
              }
            });
          }
          if (explain) explain.removeAttribute('hidden');
        });
      });
    });
  }

  /* ---------- 12. flow diagram nodes -------------------------------------- */
  function initFlow() {
    var nodes = $$('.flow-node');
    var detail = $('#flow-detail');
    if (!nodes.length || !detail) return;
    nodes.forEach(function (node) {
      node.addEventListener('click', function () {
        var open = node.getAttribute('aria-expanded') === 'true';
        nodes.forEach(function (n) { n.setAttribute('aria-expanded', 'false'); });
        if (open) { detail.setAttribute('hidden', ''); return; }
        node.setAttribute('aria-expanded', 'true');
        detail.removeAttribute('hidden');
        $('.tag', detail).textContent = node.getAttribute('data-key');
        $('h3', detail).textContent = node.getAttribute('data-title');
        $('[data-slot="what"]', detail).textContent = node.getAttribute('data-what');
        $('[data-slot="data"]', detail).textContent = node.getAttribute('data-data');
        $('[data-slot="risk"]', detail).textContent = node.getAttribute('data-risk');
      });
    });
  }

  /* ---------- 13. final go / no-go self-assessment ------------------------ */
  function initGate() {
    var form = $('#gate-form');
    var verdict = $('#gate-verdict');
    if (!form || !verdict) return;
    var boxes = $$('input[type="checkbox"]', form);

    boxes.forEach(function (b) {
      b.checked = !!state.gate[b.id];
      b.addEventListener('change', function () {
        state.gate[b.id] = b.checked; saveState(); evaluate();
      });
    });

    function evaluate() {
      var stop = boxes.filter(function (b) { return b.getAttribute('data-weight') === 'stop' && b.checked; });
      var legal = boxes.filter(function (b) { return b.getAttribute('data-weight') === 'legal' && b.checked; });
      var check = boxes.filter(function (b) { return b.getAttribute('data-weight') === 'check' && b.checked; });
      var go = boxes.filter(function (b) { return b.getAttribute('data-weight') === 'go'; });
      var goDone = go.filter(function (b) { return b.checked; });

      var cls, text, note;
      if (stop.length) {
        cls = 'status--stop'; text = 'STOP';
        note = 'Отмечен минимум один стоп-признак (' + stop.length + '). Запускать нельзя: сначала уберите причину.';
      } else if (legal.length) {
        cls = 'status--legal'; text = 'LEGAL REVIEW';
        note = 'Конфигурация выходит за рамки простого административного сценария. Нужна проверка Rechtsanwalt / Datenschutzbeauftragter до запуска.';
      } else if (check.length) {
        cls = 'status--check'; text = 'CHECK';
        note = 'Есть элементы (' + check.length + '), которые требуют документального подтверждения до пилота.';
      } else if (goDone.length === go.length && go.length) {
        cls = 'status--go'; text = 'GO (для пилота)';
        note = 'Базовые условия административной версии отмечены. Это не юридическое заключение: конфигурацию всё равно подтверждает практика и её консультант.';
      } else {
        cls = 'status--check'; text = 'НЕ ГОТОВО';
        note = 'Отмечено ' + goDone.length + ' из ' + go.length + ' обязательных условий GO.';
      }
      verdict.className = 'status ' + cls;
      verdict.textContent = text;
      $('#gate-note').textContent = note;
    }
    evaluate();
  }

  /* ---------- 14. print --------------------------------------------------- */
  function initPrint() {
    $$('[data-print]').forEach(function (btn) {
      btn.addEventListener('click', function () { window.print(); });
    });
  }

  /* ---------- 15. horizontal scroll affordance for tables ----------------- */
  function initTables() {
    $$('.table-wrap').forEach(function (w) {
      if (w.scrollWidth > w.clientWidth + 4) {
        w.setAttribute('tabindex', '0');
        w.setAttribute('role', 'region');
      }
    });
  }

  /* ---------- boot -------------------------------------------------------- */
  function boot() {
    loadState();
    initProgress();
    initTocObserver();
    initStepObserver();
    initNavToggle();
    initDisclosures();
    initFilterGroup('#risk-filters');
    initFilterGroup('#datamap-filters');
    initGlossary();
    initSteps();
    initVendorList();
    initCopy();
    initQuiz();
    initFlow();
    initGate();
    initPrint();
    initTables();
    document.documentElement.classList.add('js-ready');
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else { boot(); }
})();
