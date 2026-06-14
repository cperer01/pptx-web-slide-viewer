/* Web Slide - content add-in logic */
(function () {
  "use strict";

  var SETTING_KEY = "webSlideUrl";

  /* --- render tuning ---
     The embedded page is rendered into a fixed virtual canvas, then scaled to
     fit the slide object. This forces the app to render its full desktop
     layout regardless of how small the slide object is. */
  var CANVAS_W = 1920;
  var CANVAS_H = 1080;

  /* Pixels (in 1920x1080 space) to clip off the TOP of the embedded page,
     e.g. to hide a fixed header or demo-selector bar that lives inside the
     app. 0 = no crop. Try ~72-96 to remove a typical top button bar. */
  var CROP_TOP = 0;

  var els = {};
  var blockTimer = null;

  Office.onReady(function () {
    els.url = document.getElementById("url");
    els.reload = document.getElementById("reload");
    els.stage = document.getElementById("stage");
    els.frame = document.getElementById("frame");
    els.empty = document.getElementById("empty");
    els.warn = document.getElementById("warn");

    els.reload.addEventListener("click", onReload);
    els.url.addEventListener("keydown", function (e) {
      if (e.key === "Enter") onLoad();
    });
    els.frame.addEventListener("load", onFrameLoad);

    // Keep the canvas fitted to the slide object as PowerPoint resizes it.
    if (window.ResizeObserver) {
      new ResizeObserver(fitFrame).observe(els.stage);
    }
    window.addEventListener("resize", fitFrame);
    fitFrame();
    setTimeout(fitFrame, 60);

    var saved = readSetting();
    if (saved) {
      els.url.value = saved;
      navigate(saved);
    }
  });

  /* --- scale the 1920x1080 canvas to fit the slide object, centered --- */
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
    writeSetting(url);
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
  function readSetting() {
    try {
      return Office.context.document.settings.get(SETTING_KEY) || "";
    } catch (e) {
      return "";
    }
  }

  function writeSetting(url) {
    try {
      Office.context.document.settings.set(SETTING_KEY, url);
      Office.context.document.settings.saveAsync();
    } catch (e) {
      /* settings unavailable in some contexts; ignore */
    }
  }

  /* --- tiny DOM helpers --- */
  function show(el) { el.classList.remove("hidden"); }
  function hide(el) { el.classList.add("hidden"); }
})();
