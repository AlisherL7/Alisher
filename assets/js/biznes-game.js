(function () {
  const T = window.GAME_I18N;
  if (!T) return;

  const MAX_WEEKS = 12;
  const TEA_PRICE_REF = 3;
  const SAMSA_PRICE_REF = 6;
  const TEA_DEMAND_REF = 80;
  const SAMSA_DEMAND_REF = 40;
  const ELASTICITY = 1.4;
  const STORAGE_KEY = 'biznesGameBest';

  const EVENT_KEYS = Object.keys(T.events);

  let state;
  let inputs;

  function newState() {
    return {
      week: 1,
      cash: 500,
      reputation: 50,
      demandMultPersist: 1,
      log: [],
      over: false,
    };
  }

  function newInputs() {
    return {
      teaPrice: 3,
      samsaPrice: 6,
      teaStock: 70,
      samsaStock: 30,
      staff: 1,
    };
  }

  function pickEvent() {
    if (Math.random() < 0.35) return null;
    return EVENT_KEYS[Math.floor(Math.random() * EVENT_KEYS.length)];
  }

  function applyEventEffects(eventKey) {
    const effects = { demandMult: 1, costMult: 1, stockPenalty: 1, extraFixedCost: 0 };
    switch (eventKey) {
      case 'heatwave':
        effects.demandMult = 1.35;
        break;
      case 'holiday':
        effects.demandMult = 1.5;
        break;
      case 'competitor':
        state.demandMultPersist *= 0.85;
        break;
      case 'inflation':
        effects.costMult = 1.3;
        break;
      case 'fx':
        effects.costMult = 1.2;
        break;
      case 'tax':
        effects.extraFixedCost = 40;
        break;
      case 'goodpress':
        state.reputation = Math.min(100, state.reputation + 12);
        break;
      case 'shortage':
        effects.stockPenalty = 0.7;
        break;
    }
    return effects;
  }

  function computeWeek(effects) {
    const { teaPrice, samsaPrice, teaStock, samsaStock, staff } = inputs;

    const repFactor = 0.5 + state.reputation / 100;
    const demandMult = effects.demandMult * state.demandMultPersist;

    const teaDemand = TEA_DEMAND_REF * Math.pow(TEA_PRICE_REF / teaPrice, ELASTICITY) * demandMult * repFactor;
    const samsaDemand = SAMSA_DEMAND_REF * Math.pow(SAMSA_PRICE_REF / samsaPrice, ELASTICITY) * demandMult * repFactor;

    const effTeaStock = teaStock * effects.stockPenalty;
    const effSamsaStock = samsaStock * effects.stockPenalty;

    const capacity = staff * 70;

    const teaSold = Math.min(teaDemand, effTeaStock, capacity);
    const samsaSold = Math.min(samsaDemand, effSamsaStock, Math.max(0, capacity - teaSold));

    const revenue = teaSold * teaPrice + samsaSold * samsaPrice;
    const ingredientCost = (teaStock * 1 + samsaStock * 3) * effects.costMult;
    const wages = (staff - 1) * 40;
    const rent = 50 + effects.extraFixedCost;
    const totalCost = ingredientCost + wages + rent;
    const profit = revenue - totalCost;

    let repChange = 0;
    if (teaDemand > effTeaStock * 1.1 || samsaDemand > effSamsaStock * 1.1) repChange -= 5;
    if (profit > 0) repChange += 2;

    state.reputation = Math.max(0, Math.min(100, state.reputation + repChange));
    state.cash += profit;

    return { teaSold, samsaSold, revenue, ingredientCost, wages, rent, totalCost, profit };
  }

  function getRank(cash) {
    if (cash < 0) return T.ranks.bankrupt;
    if (cash < 500) return T.ranks.beginner;
    if (cash < 1200) return T.ranks.stable;
    if (cash < 2500) return T.ranks.successful;
    return T.ranks.magnate;
  }

  function fmt(n) {
    return Math.round(n).toLocaleString('ru-RU');
  }

  document.addEventListener('DOMContentLoaded', () => {
    const els = {
      week: document.getElementById('stat-week'),
      cash: document.getElementById('stat-cash'),
      rep: document.getElementById('stat-rep'),
      banner: document.getElementById('game-event-banner'),
      log: document.getElementById('game-log'),
      end: document.getElementById('game-end'),
      nextBtn: document.getElementById('next-week-btn'),
      controls: document.getElementById('game-controls-panel'),

      teaPrice: document.getElementById('ctrl-tea-price'),
      teaPriceVal: document.getElementById('val-tea-price'),
      samsaPrice: document.getElementById('ctrl-samsa-price'),
      samsaPriceVal: document.getElementById('val-samsa-price'),
      teaStock: document.getElementById('ctrl-tea-stock'),
      teaStockVal: document.getElementById('val-tea-stock'),
      samsaStock: document.getElementById('ctrl-samsa-stock'),
      samsaStockVal: document.getElementById('val-samsa-stock'),
      staff: document.getElementById('ctrl-staff'),
      staffVal: document.getElementById('val-staff'),
    };

    function bindRange(input, valueEl, suffix, isInt) {
      const update = () => {
        const v = parseFloat(input.value);
        inputs[input.dataset.field] = v;
        valueEl.textContent = (isInt ? Math.round(v) : v) + suffix;
      };
      input.addEventListener('input', update);
      update();
    }

    function setupControls() {
      els.teaPrice.value = inputs.teaPrice;
      els.samsaPrice.value = inputs.samsaPrice;
      els.teaStock.value = inputs.teaStock;
      els.samsaStock.value = inputs.samsaStock;
      els.staff.value = inputs.staff;

      bindRange(els.teaPrice, els.teaPriceVal, ' ' + T.ui.currency, false);
      bindRange(els.samsaPrice, els.samsaPriceVal, ' ' + T.ui.currency, false);
      bindRange(els.teaStock, els.teaStockVal, ' ' + T.ui.cupsUnit, true);
      bindRange(els.samsaStock, els.samsaStockVal, ' ' + T.ui.piecesUnit, true);
      bindRange(els.staff, els.staffVal, ' ' + T.ui.staffUnit, true);
    }

    function updateStats() {
      els.week.textContent = state.week + ' / ' + MAX_WEEKS;
      els.cash.textContent = fmt(state.cash) + ' ' + T.ui.currency;
      els.rep.textContent = Math.round(state.reputation) + '%';
    }

    function showEvent(eventKey, effects) {
      if (!eventKey) {
        els.banner.className = 'game-event';
        els.banner.style.display = '';
        els.banner.innerHTML = T.ui.quietWeek;
        return;
      }
      const ev = T.events[eventKey];
      els.banner.className = 'game-event ' + ev.type;
      els.banner.style.display = '';
      const link = `<a href="${ev.lessonPage}#lesson-${ev.lessonAnchor}">${T.ui.relatedLessonLabel}: ${ev.lessonTitle} →</a>`;
      els.banner.innerHTML = `<div>${ev.text}</div><div style="margin-top:8px;font-size:0.85rem;">💡 ${link}</div>`;
    }

    function addLogEntry(weekNumber, result) {
      const entry = document.createElement('div');
      entry.className = 'game-log-entry';
      const profitClass = result.profit >= 0 ? 'profit' : 'loss';
      const sign = result.profit >= 0 ? '+' : '';
      entry.innerHTML = `
        <span>${T.ui.weekLabel} ${weekNumber}: ${T.ui.soldLabel} ${Math.round(result.teaSold)} ${T.ui.cupsUnit} + ${Math.round(result.samsaSold)} ${T.ui.piecesUnit}</span>
        <span class="${profitClass}">${sign}${fmt(result.profit)} ${T.ui.currency}</span>
      `;
      els.log.prepend(entry);
    }

    function endGame() {
      state.over = true;
      els.controls.style.display = 'none';
      els.nextBtn.style.display = 'none';
      els.end.style.display = 'block';

      const rank = getRank(state.cash);
      const best = parseFloat(localStorage.getItem(STORAGE_KEY) || '-Infinity');
      let recordHtml = '';
      if (state.cash > best) {
        localStorage.setItem(STORAGE_KEY, String(state.cash));
        recordHtml = `<p class="game-record">🏆 ${T.ui.newRecord}</p>`;
      } else if (isFinite(best)) {
        recordHtml = `<p>${T.ui.bestScoreLabel}: ${fmt(best)} ${T.ui.currency}</p>`;
      }

      const title = state.cash < 0 ? T.ui.gameOverBankrupt : T.ui.gameOverFinished;

      els.end.innerHTML = `
        <h3>${title}</h3>
        <div class="game-rank">${rank}</div>
        <p>${T.ui.finalCashLabel}: <strong>${fmt(state.cash)} ${T.ui.currency}</strong></p>
        ${recordHtml}
        <p class="game-end-lessons">${T.ui.lessonsHint}</p>
        <button id="restart-btn-2" class="btn btn-secondary">${T.ui.restartBtn}</button>
      `;
      document.getElementById('restart-btn-2').addEventListener('click', restart);
    }

    function nextWeek() {
      if (state.over) return;
      const eventKey = pickEvent();
      const effects = applyEventEffects(eventKey);
      const result = computeWeek(effects);

      addLogEntry(state.week, result);
      showEvent(eventKey, effects);

      state.week += 1;
      updateStats();

      if (state.cash < 0 || state.week > MAX_WEEKS) {
        endGame();
      }
    }

    function restart() {
      state = newState();
      inputs = newInputs();
      els.log.innerHTML = '';
      els.banner.style.display = 'none';
      els.controls.style.display = '';
      els.nextBtn.style.display = '';
      els.end.style.display = 'none';
      setupControls();
      updateStats();
    }

    state = newState();
    inputs = newInputs();
    setupControls();
    updateStats();

    els.nextBtn.addEventListener('click', nextWeek);
  });
})();
