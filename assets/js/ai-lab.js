// ===================== ИИ-Лаборатория инвестиций =====================
// Простой, но настоящий алгоритмический бэктестер: SMA-пересечение + фильтр RSI.
// Если живые данные недоступны (нет сети / лимит API), используется
// детерминированный реалистичный набор данных, чтобы демо всегда работало.

(function () {
  "use strict";

  var T = window.AI_LAB_I18N || {};

  // ---------- Детерминированный генератор "реалистичного" рынка ----------
  function mulberry32(seed) {
    return function () {
      seed |= 0;
      seed = (seed + 0x6d2b79f5) | 0;
      var t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  function generateSeries(seed, days, startPrice, drift, volatility) {
    var rng = mulberry32(seed);
    var prices = [startPrice];
    for (var i = 1; i < days; i++) {
      var r = (rng() + rng() + rng() - 1.5) / 1.5; // ~ -1..1, ближе к нормальному
      var change = drift + volatility * r;
      var next = prices[i - 1] * (1 + change);
      if (next < 0.01) next = 0.01;
      prices.push(next);
    }
    return prices;
  }

  var ASSET_CONFIG = {
    btc: { seed: 138, days: 240, startPrice: 42000, drift: 0.0016, volatility: 0.035, coingeckoId: "bitcoin", decimals: 0 },
    eth: { seed: 138, days: 240, startPrice: 2200, drift: 0.0013, volatility: 0.045, coingeckoId: "ethereum", decimals: 2 },
    stock: { seed: 138, days: 240, startPrice: 150, drift: 0.0009, volatility: 0.016, coingeckoId: null, decimals: 2 }
  };

  var dataCache = {};

  function getSeries(key) {
    if (!dataCache[key]) {
      var cfg = ASSET_CONFIG[key];
      dataCache[key] = {
        prices: generateSeries(cfg.seed, cfg.days, cfg.startPrice, cfg.drift, cfg.volatility),
        live: false
      };
    }
    return dataCache[key];
  }

  // ---------- Индикаторы ----------
  function sma(data, period) {
    var out = new Array(data.length).fill(null);
    for (var i = period - 1; i < data.length; i++) {
      var sum = 0;
      for (var j = i - period + 1; j <= i; j++) sum += data[j];
      out[i] = sum / period;
    }
    return out;
  }

  function rsi(data, period) {
    var out = new Array(data.length).fill(null);
    for (var i = period; i < data.length; i++) {
      var gains = 0, losses = 0;
      for (var j = i - period + 1; j <= i; j++) {
        var diff = data[j] - data[j - 1];
        if (diff >= 0) gains += diff; else losses -= diff;
      }
      var avgGain = gains / period;
      var avgLoss = losses / period;
      if (avgLoss === 0) {
        out[i] = 100;
      } else {
        var rs = avgGain / avgLoss;
        out[i] = 100 - 100 / (1 + rs);
      }
    }
    return out;
  }

  // ---------- Бэктест: SMA-пересечение + фильтр RSI ----------
  function backtest(prices, params) {
    var shortSMA = sma(prices, params.shortPeriod);
    var longSMA = sma(prices, params.longPeriod);
    var rsiVals = rsi(prices, params.rsiPeriod);

    var n = prices.length;
    var position = 0; // 0 = кэш, 1 = в позиции
    var entryPrice = null;
    var trades = [];
    var equity = new Array(n).fill(100);
    var buyHold = new Array(n).fill(100);
    var markers = [];
    var daysInMarket = 0;

    for (var i = 1; i < n; i++) {
      buyHold[i] = 100 * (prices[i] / prices[0]);

      if (position === 1) {
        equity[i] = equity[i - 1] * (prices[i] / prices[i - 1]);
        daysInMarket++;
      } else {
        equity[i] = equity[i - 1];
      }

      if (
        shortSMA[i] == null || longSMA[i] == null || rsiVals[i] == null ||
        shortSMA[i - 1] == null || longSMA[i - 1] == null
      ) {
        continue;
      }

      var crossUp = shortSMA[i] > longSMA[i] && shortSMA[i - 1] <= longSMA[i - 1];
      var crossDown = shortSMA[i] < longSMA[i] && shortSMA[i - 1] >= longSMA[i - 1];
      var overbought = rsiVals[i] > params.rsiThreshold;

      if (position === 0 && crossUp && !overbought) {
        position = 1;
        entryPrice = prices[i];
        markers.push({ index: i, type: "buy" });
      } else if (position === 1 && (crossDown || rsiVals[i] > params.rsiThreshold + 10)) {
        position = 0;
        trades.push({ entry: entryPrice, exit: prices[i], pnlPct: (prices[i] - entryPrice) / entryPrice * 100 });
        markers.push({ index: i, type: "sell" });
        entryPrice = null;
      }
    }

    if (position === 1 && entryPrice != null) {
      trades.push({ entry: entryPrice, exit: prices[n - 1], pnlPct: (prices[n - 1] - entryPrice) / entryPrice * 100, open: true });
    }

    var peak = equity[0], maxDD = 0;
    for (var k = 0; k < n; k++) {
      if (equity[k] > peak) peak = equity[k];
      var dd = (peak - equity[k]) / peak * 100;
      if (dd > maxDD) maxDD = dd;
    }

    var wins = trades.filter(function (t) { return t.pnlPct > 0; }).length;

    return {
      prices: prices,
      shortSMA: shortSMA,
      longSMA: longSMA,
      rsiVals: rsiVals,
      equity: equity,
      buyHold: buyHold,
      markers: markers,
      trades: trades,
      stats: {
        totalReturn: equity[n - 1] - 100,
        buyHoldReturn: buyHold[n - 1] - 100,
        maxDrawdown: maxDD,
        tradesCount: trades.length,
        winRate: trades.length ? (wins / trades.length * 100) : 0,
        timeInMarket: (daysInMarket / (n - 1)) * 100
      }
    };
  }

  // ---------- Рисование графиков на canvas ----------
  function drawLineChart(canvas, series, opts) {
    opts = opts || {};
    var dpr = window.devicePixelRatio || 1;
    var rect = canvas.getBoundingClientRect();
    var W = Math.max(rect.width, 240);
    var H = canvas.clientHeight || 260;
    canvas.width = W * dpr;
    canvas.height = H * dpr;
    var ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, W, H);

    var padding = { top: 14, right: 14, bottom: 22, left: 56 };
    var plotW = W - padding.left - padding.right;
    var plotH = H - padding.top - padding.bottom;

    var min = Infinity, max = -Infinity;
    series.forEach(function (s) {
      s.data.forEach(function (v) {
        if (v == null) return;
        if (v < min) min = v;
        if (v > max) max = v;
      });
    });
    if (!isFinite(min) || !isFinite(max)) return;
    if (min === max) { min -= 1; max += 1; }
    var pad = (max - min) * 0.08;
    min -= pad;
    max += pad;

    var n = series[0].data.length;
    var xStep = plotW / (n - 1);

    function xPos(i) { return padding.left + i * xStep; }
    function yPos(v) { return padding.top + plotH - ((v - min) / (max - min)) * plotH; }

    ctx.strokeStyle = "#e3e8f0";
    ctx.fillStyle = "#6b7280";
    ctx.font = "11px 'Noto Sans', sans-serif";
    ctx.lineWidth = 1;
    var gridLines = 4;
    for (var g = 0; g <= gridLines; g++) {
      var v = min + (max - min) * (g / gridLines);
      var y = yPos(v);
      ctx.beginPath();
      ctx.moveTo(padding.left, y);
      ctx.lineTo(W - padding.right, y);
      ctx.stroke();
      var label = opts.formatY ? opts.formatY(v) : v.toFixed(1);
      ctx.fillText(label, 4, y + 4);
    }

    series.forEach(function (s) {
      ctx.beginPath();
      ctx.strokeStyle = s.color;
      ctx.lineWidth = s.width || 2;
      var started = false;
      for (var i = 0; i < n; i++) {
        var val = s.data[i];
        if (val == null) { started = false; continue; }
        var x = xPos(i), yv = yPos(val);
        if (!started) { ctx.moveTo(x, yv); started = true; }
        else ctx.lineTo(x, yv);
      }
      ctx.stroke();
    });

    if (opts.markers && opts.markerSeries) {
      opts.markers.forEach(function (m) {
        var price = opts.markerSeries[m.index];
        if (price == null) return;
        var x = xPos(m.index), y = yPos(price);
        ctx.beginPath();
        ctx.fillStyle = m.type === "buy" ? "#15803d" : "#b91c1c";
        if (m.type === "buy") {
          ctx.moveTo(x, y + 8);
          ctx.lineTo(x - 6, y + 19);
          ctx.lineTo(x + 6, y + 19);
        } else {
          ctx.moveTo(x, y - 8);
          ctx.lineTo(x - 6, y - 19);
          ctx.lineTo(x + 6, y - 19);
        }
        ctx.closePath();
        ctx.fill();
      });
    }
  }

  // ---------- UI ----------
  document.addEventListener("DOMContentLoaded", function () {
    var assetTabs = document.getElementById("asset-tabs");
    if (!assetTabs) return; // не страница ИИ-лаборатории

    var priceChart = document.getElementById("price-chart");
    var equityChart = document.getElementById("equity-chart");
    var priceChartSub = document.getElementById("price-chart-sub");
    var statGrid = document.getElementById("stat-grid");
    var liveStatus = document.getElementById("live-status");
    var liveBtn = document.getElementById("live-btn");

    var shortInput = document.getElementById("short-period");
    var longInput = document.getElementById("long-period");
    var rsiPeriodInput = document.getElementById("rsi-period");
    var rsiThresholdInput = document.getElementById("rsi-threshold");

    var valShort = document.getElementById("val-short");
    var valLong = document.getElementById("val-long");
    var valRsi = document.getElementById("val-rsi");
    var valRsiTh = document.getElementById("val-rsi-th");

    var currentAsset = "btc";

    var STAT_DEFS = [
      { key: "totalReturn", label: T.statTotalReturn, unit: "%", color: "auto" },
      { key: "buyHoldReturn", label: T.statBuyHold, unit: "%", color: "auto" },
      { key: "maxDrawdown", label: T.statMaxDrawdown, unit: "%", color: "negative" },
      { key: "tradesCount", label: T.statTrades, unit: "", color: "neutral" },
      { key: "winRate", label: T.statWinRate, unit: "%", color: "auto" },
      { key: "timeInMarket", label: T.statTimeInMarket, unit: "%", color: "neutral" }
    ];

    function formatPrice(value, key) {
      var decimals = ASSET_CONFIG[key].decimals;
      return "$" + value.toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
    }

    function updateAssetTabPrices() {
      Object.keys(ASSET_CONFIG).forEach(function (key) {
        var el = document.getElementById("price-" + key);
        if (!el) return;
        var series = getSeries(key);
        var last = series.prices[series.prices.length - 1];
        el.textContent = formatPrice(last, key);
      });
    }

    function readParams() {
      return {
        shortPeriod: parseInt(shortInput.value, 10),
        longPeriod: parseInt(longInput.value, 10),
        rsiPeriod: parseInt(rsiPeriodInput.value, 10),
        rsiThreshold: parseInt(rsiThresholdInput.value, 10)
      };
    }

    function syncLabels() {
      valShort.textContent = shortInput.value;
      valLong.textContent = longInput.value;
      valRsi.textContent = rsiPeriodInput.value;
      valRsiTh.textContent = rsiThresholdInput.value;
    }

    function render() {
      syncLabels();
      var params = readParams();
      if (params.shortPeriod >= params.longPeriod) {
        params.longPeriod = params.shortPeriod + 5;
      }

      var dataset = getSeries(currentAsset);
      var result = backtest(dataset.prices, params);
      var cfg = ASSET_CONFIG[currentAsset];

      priceChartSub.textContent = dataset.live
        ? (T.chartSubLive || "")
        : (T.chartSubDemo || "");

      drawLineChart(priceChart, [
        { data: result.prices, color: "#1450a3", width: 2 },
        { data: result.shortSMA, color: "#f4a93b", width: 1.5 },
        { data: result.longSMA, color: "#b91c1c", width: 1.5 }
      ], {
        formatY: function (v) { return formatPrice(v, currentAsset); },
        markers: result.markers,
        markerSeries: result.prices
      });

      drawLineChart(equityChart, [
        { data: result.equity, color: "#15803d", width: 2 },
        { data: result.buyHold, color: "#6b7280", width: 2 }
      ], {
        formatY: function (v) { return v.toFixed(0) + "%"; }
      });

      statGrid.innerHTML = "";
      STAT_DEFS.forEach(function (def) {
        var raw = result.stats[def.key];
        var valueStr;
        if (def.unit === "%") {
          valueStr = (raw > 0 && def.key !== "maxDrawdown" ? "+" : "") + raw.toFixed(1) + "%";
        } else {
          valueStr = String(Math.round(raw));
        }

        var colorClass = "";
        if (def.color === "negative") colorClass = "negative";
        else if (def.color === "auto") colorClass = raw >= 0 ? "positive" : "negative";

        var card = document.createElement("div");
        card.className = "stat-card";
        card.innerHTML =
          '<div class="stat-label">' + (def.label || def.key) + '</div>' +
          '<div class="stat-value ' + colorClass + '">' + valueStr + '</div>';
        statGrid.appendChild(card);
      });
    }

    // ---------- Переключение активов ----------
    assetTabs.querySelectorAll("button[data-asset]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        assetTabs.querySelectorAll("button[data-asset]").forEach(function (b) { b.classList.remove("active"); });
        btn.classList.add("active");
        currentAsset = btn.getAttribute("data-asset");
        liveStatus.textContent = "";
        liveStatus.className = "live-data-status";
        render();
      });
    });

    // ---------- Контролы ----------
    [shortInput, longInput, rsiPeriodInput, rsiThresholdInput].forEach(function (input) {
      input.addEventListener("input", syncLabels);
    });

    var runBtn = document.getElementById("run-btn");
    runBtn.addEventListener("click", render);

    // ---------- Живые данные (CoinGecko) ----------
    if (liveBtn) {
      liveBtn.addEventListener("click", function () {
        var cfg = ASSET_CONFIG[currentAsset];
        if (!cfg.coingeckoId) {
          liveStatus.textContent = T.liveUnsupported || "";
          liveStatus.className = "live-data-status error";
          return;
        }

        liveStatus.textContent = T.liveLoading || "...";
        liveStatus.className = "live-data-status";

        var url = "https://api.coingecko.com/api/v3/coins/" + cfg.coingeckoId +
          "/market_chart?vs_currency=usd&days=240&interval=daily";

        fetch(url)
          .then(function (resp) {
            if (!resp.ok) throw new Error("bad response");
            return resp.json();
          })
          .then(function (data) {
            if (!data.prices || !data.prices.length) throw new Error("no data");
            var prices = data.prices.map(function (p) { return p[1]; });
            dataCache[currentAsset] = { prices: prices, live: true };
            liveStatus.textContent = T.liveSuccess || "";
            liveStatus.className = "live-data-status success";
            updateAssetTabPrices();
            render();
          })
          .catch(function () {
            liveStatus.textContent = T.liveError || "";
            liveStatus.className = "live-data-status error";
          });
      });
    }

    // ---------- Resize ----------
    var resizeTimer = null;
    window.addEventListener("resize", function () {
      if (resizeTimer) clearTimeout(resizeTimer);
      resizeTimer = setTimeout(render, 150);
    });

    updateAssetTabPrices();
    render();
  });
})();
