/* Web Slide - content add-in logic */
(function () {
  "use strict";

  /* --- per-slide persisted settings ---
     Each value is stored in the slide's own add-in settings, so reopening the
     deck restores the URL, the chosen resolution, and whether the bar is hidden. */
  var URL_KEY = "webSlideUrl";
  var RES_KEY = "webSlideRes";
  var BAR_KEY = "webSlideBarHidden";

  /* --- render tuning ---
     The embedded page is rendered into a virtual canvas, then scaled to fit the
     slide object. This forces the app to render its full desktop layout
     regardless of how small the slide object is. The canvas size (the
     "resolution") is chosen from the selector in the bar and saved per slide. */
  var CANVAS_W = 1920;
  var CANVAS_H = 1080;

  /* Common canvas resolutions offered in the selector (set while editing). */
  var RESOLUTIONS = [
    { w: 1280, h: 720,  label: "1280 × 720 (720p)" },
    { w: 1366, h: 768,  label: "1366 × 768" },
    { w: 1440, h: 900,  label: "1440 × 900" },
    { w: 1600, h: 900,  label: "1600 × 900" },
    { w: 1920, h: 1080, label: "1920 × 1080 (1080p)" },
    { w: 2560, h: 1440, label: "2560 × 1440 (1440p)" },
    { w: 3840, h: 2160, label: "3840 × 2160 (4K)" },
    { w: 1024, h: 768,  label: "1024 × 768 (4:3)" },
    { w: 1280, h: 1024, label: "1280 × 1024 (5:4)" }
  ];
  var DEFAULT_RES = "1920x1080";

  /* Pixels (in canvas space) to clip off the TOP of the embedded page, e.g. to
     hide a fixed header that lives inside the app. 0 = no crop. */
  var CROP_TOP = 0;

  var els = {};
  var blockTimer = null;

  Office.onReady(function () {
    els.app = document.getElementById("app");
    els.bar = document.getElementById("bar");
    els.res = document.getElementById("res");
    els.url = document.getElementById("url");
    els.reload = document.getElementById("reload");
    els.hide = document.getElementById("hide");
    els.show = document.getElementById("show");
    els.stage = document.getElementById("stage");
    els.frame = document.getElementById("frame");
    els.empty = document.getElementById("empty");
    els.warn = document.getElementById("warn");

    buildResOptions();

    els.reload.addEventListener("click", onReload);
    els.url.addEventListener("keydown", function (e) {
      if (e.key === "Enter") onLoad();
    });
    els.res.addEventListener("change", onResChange);
    els.hide.addEventListener("click", function () { setBarHidden(true); });
    els.show.addEventListener("click", function () { setBarHidden(false); });
    els.frame.addEventListener("load", onFrameLoad);

    // Keep the canvas fitted to the slide object as PowerPoint resizes it.
    if (window.ResizeObserver) {
      new ResizeObserver(fitFrame).observe(els.stage);
    }
    window.addEventListener("resize", fitFrame);

    // Restore saved state: resolution and bar visibility first, then the URL.
    applyResolution(getSetting(RES_KEY) || DEFAULT_RES);
    setBarHidden(getSetting(BAR_KEY) === true, true);

    fitFrame();
    setTimeout(fitFrame, 60);

    var saved = getSetting(URL_KEY);
    if (saved) {
      els.url.value = saved;
      navigate(saved);
    }
  });

  /* --- resolution selector (edit-time) --- */
  function buildResOptions() {
    for (var i = 0; i < RESOLUTIONS.length; i++) {
      var r = RESOLUTIONS[i];
      var opt = document.createElement("option");
      opt.value = r.w + "x" + r.h;
      opt.textContent = r.label;
      els.res.appendChild(opt);
    }
  }

  function onResChange() {
    applyResolution(els.res.value);
    setSetting(RES_KEY, els.res.value);
  }

  function applyResolution(value) {
    var parts = /^(\d+)x(\d+)$/.exec(value || "");
    if (!parts) {
      value = DEFAULT_RES;
      parts = /^(\d+)x(\d+)$/.exec(value);
    }
    CANVAS_W = parseInt(parts[1], 10);
    CANVAS_H = parseInt(parts[2], 10);
    els.res.value = value;
    els.frame.style.width = CANVAS_W + "px";
    els.frame.style.height = CANVAS_H + "px";
    fitFrame();
  }

  /* --- URL / reload bar show-hide (works in edit and slideshow) --- */
  function setBarHidden(hidden, skipSave) {
    els.app.classList.toggle("bar-hidden", !!hidden);
    if (!skipSave) setSetting(BAR_KEY, !!hidden);
    fitFrame();
  }

  /* --- scale the virtual canvas to fit the slide object, centered --- */
  function fitFrame() {
    var w = els.stage.clientWidth;
    var h = els.stage.clientHeight;
    if (!w || !h) return;

    var visibleH = CANVAS_H - CROP_TOP;
    var s = Math.min(w / CANVAS_W, h / visibleH);

    var scaledW = CANVAS_W * s;
    var scaledBandH = visibleH * s;
    var offsetX = (w - scaledW) / 2;
    var offsetY = (h - scaledBandH) / 2;

    // CSS applies scale first, then translate (in final pixels). Shift up by
    // the cropped band so the visible region starts below it.
    var tx = offsetX;
    var ty = offsetY - CROP_TOP * s;

    els.frame.style.transform =
      "translate(" + tx + "px, " + ty + "px) scale(" + s + ")";
  }

  function normalize(raw) {
    var v = (raw || "").trim();
    if (!v) return "";
    if (!/^https?:\/\//i.test(v)) v = "https://" + v;
    return v;
  }

  function onLoad() {
    var url = normalize(els.url.value);
    if (!url) return;
    els.url.value = url;
    navigate(url);
    setSetting(URL_KEY, url);
  }

  function onReload() {
    var url = normalize(els.url.value);
    if (!url) return;
    els.frame.src = "about:blank";
    setTimeout(function () { navigate(url); }, 50);
  }

  function navigate(url) {
    hide(els.empty);
    hide(els.warn);
    if (blockTimer) clearTimeout(blockTimer);

    els.frame.src = url;
    fitFrame();

    blockTimer = setTimeout(function () {
      show(els.warn);
    }, 4000);
  }

  function onFrameLoad() {
    if (els.frame.src && els.frame.src !== "about:blank") {
      if (blockTimer) clearTimeout(blockTimer);
      hide(els.warn);
      fitFrame();
    }
  }

  /* --- persistence via the slide's add-in settings --- */
  function getSetting(key) {
    try {
      return Office.context.document.settings.get(key);
    } catch (e) {
      return null;
    }
  }

  function setSetting(key, val) {
    try {
      Office.context.document.settings.set(key, val);
      Office.context.document.settings.saveAsync();
    } catch (e) {
      /* settings unavailable in some contexts; ignore */
    }
  }

  /* --- tiny DOM helpers --- */
  function show(el) { el.classList.remove("hidden"); }
  function hide(el) { el.classList.add("hidden"); }
})();
