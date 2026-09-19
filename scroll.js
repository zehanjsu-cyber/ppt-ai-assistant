// Automatic answer scrolling stays inside the add-in, not the PowerPoint slide.
(() => {
  const SETTINGS_KEY = "ppt-ai-scroll-options";
  let answer, enabledInput, secondsInput, hint;
  let ownerSlide = null, visibleSlide = null, timer = null;
  let checking = false, generation = 0, ready = false, done = false, phase = "idle";
  let progress;
  let fallback = null, fallbackHeight = 0, fallbackTop = 0;
  let misses = 0;
  const position = () => fallback ? fallbackTop : answer.scrollTop;
  const height = () => fallback ? fallbackHeight : answer.scrollHeight;
  function moveFallback(top) {
    fallbackTop = top;
    fallback.style.transform = "translateY(-" + top + "px)";
  }
  function resetPosition() {
    if (fallback) {
      const text = fallback.textContent;
      fallback = null; fallbackHeight = 0; fallbackTop = 0;
      answer.textContent = text;
      answer.style.overflowY = "auto";
    }
    answer.scrollTop = 0;
  }
  function moveTo(top) {
    const target = Math.max(0, Math.min(height() - answer.clientHeight, top));
    if (fallback) return moveFallback(target);
    answer.scrollTop = target;
    if (Math.abs(answer.scrollTop - target) > 2) {
      fallbackHeight = answer.scrollHeight;
      fallback = document.createElement("div");
      fallback.textContent = answer.textContent;
      answer.replaceChildren(fallback);
      answer.style.overflowY = "hidden"; answer.scrollTop = 0;
      moveFallback(target);
    }
  }
  function reveal(offset, reset) {
    if (!ready || visibleSlide !== ownerSlide) return;
    if (reset) moveTo(0);
    try {
      const textNode = fallback ? fallback.firstChild : answer.firstChild;
      if (!textNode || textNode.nodeType !== 3) return;
      const range = document.createRange();
      const start = Math.max(0, Math.min(offset, textNode.length - 1));
      range.setStart(textNode, start); range.setEnd(textNode, start + 1);
      const rect = range.getBoundingClientRect(), viewport = answer.getBoundingClientRect();
      if (rect.top < viewport.top || rect.bottom > viewport.bottom - 20) {
        moveTo(position() + rect.top - viewport.top - 8);
      }
      progress.textContent = "语音同步 · 跟随正在朗读的内容";
    } catch (_) { /* Keep speech running if host geometry is unavailable. */ }
  }

  function options() {
    const saved = Office.context.document.settings.get(SETTINGS_KEY) || {};
    return {
      enabled: saved.enabled !== false,
      seconds: Number.isInteger(saved.seconds) && saved.seconds >= 5 && saved.seconds <= 120
        ? saved.seconds : 15,
    };
  }

  function stop() {
    generation++;
    clearTimeout(timer);
    timer = null;
    phase = "idle";
    if (answer) answer.scrollTo({ top: answer.scrollTop, behavior: "instant" });
  }

  function schedule() {
    if (phase !== "idle" || !ready || done || window.pptVoiceReading || !options().enabled || !visibleSlide || visibleSlide !== ownerSlide || answer.clientHeight <= 0) return;
    const atEnd = height() <= answer.clientHeight + position() + 2;
    progress.textContent = atEnd ? "最后一屏 · 阅读后停止" : "完整解析 · 每屏停留 " + options().seconds + " 秒";
    const token = generation;
    phase = "dwell";
    timer = setTimeout(async () => {
      phase = "checking";
      const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
      if (token !== generation) return;
      if (view !== "read" || !slide || slide.id !== ownerSlide || !options().enabled || window.pptVoiceReading) { stop(); return; }
      if (height() <= answer.clientHeight + position() + 2) {
        done = true; timer = null; phase = "idle";
        progress.textContent = "全部解析已展示";
        return;
      }
      phase = "moving";
      const overlap = Math.min(60, answer.clientHeight * 0.15);
      const before = position();
      const target = Math.min(height() - answer.clientHeight,
          position() + Math.max(1, answer.clientHeight - overlap)),
        movementToken = generation;
      if (fallback) {
        moveFallback(target); phase = "idle"; timer = null; schedule(); return;
      }
      answer.scrollTo({
        top: target,
        behavior: "smooth",
      });
      // Start a fresh dwell only after movement has settled, not after a fixed delay.
      let previous = -1, stable = 0, attempts = 0;
      function settle() {
        if (token !== generation) return;
        stable = Math.abs(answer.scrollTop - previous) < .5 ? stable + 1 : 0;
        previous = answer.scrollTop;
        if (stable >= 4 || ++attempts >= 20) {
          // Some embedded WebViews ignore smooth scrolling: retry directly.
          if (target > before + 2 && answer.scrollTop <= before + 2 && movementToken === generation) {
            answer.scrollTop = target;
            if (answer.scrollTop <= before + 2) {
              // If the host blocks both forms of scrolling, move text itself.
              fallbackHeight = answer.scrollHeight;
              fallback = document.createElement("div");
              fallback.textContent = answer.textContent;
              answer.replaceChildren(fallback);
              answer.style.overflowY = "hidden";
              answer.scrollTop = 0;
              moveFallback(target);
              progress.textContent = "已切换自动分屏展示";
            }
          }
          timer = null; phase = "idle"; schedule();
        }
        else timer = setTimeout(settle, 100);
      }
      timer = setTimeout(settle, 100);
    }, options().seconds * 1000);
  }

  async function check() {
    if (checking) return;
    checking = true;
    try {
      const token = generation;
      const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
      if (token !== generation) return;
      // Presenter View may background this WebView while projecting it.
      // Office's view/slide, not browser focus or visibility, is authoritative.
      if (!view || (view === "read" && !slide)) {
        // A transient Office read failure must not restart the dwell or reset text.
        if (++misses >= 4) stop();
        return;
      }
      misses = 0;
      const current = view === "read" && slide ? slide.id : null;
      document.body.classList.toggle("show-view", view === "read");
      if (current !== visibleSlide) {
        stop();
        visibleSlide = current;
        resetPosition();
        done = false;
      }
      schedule();
    } finally {
      checking = false;
    }
  }

  Office.onReady((info) => {
    if (info.host !== Office.HostType.PowerPoint) return;
    answer = document.getElementById("answer");
    enabledInput = document.getElementById("autoScroll");
    secondsInput = document.getElementById("scrollSeconds");
    hint = document.getElementById("scrollStatus");
    progress = document.getElementById("readingStatus");
    enabledInput.checked = options().enabled;
    secondsInput.value = options().seconds;
    document.getElementById("saveScroll").addEventListener("click", () => {
      const seconds = Number(secondsInput.value);
      if (!Number.isInteger(seconds) || seconds < 5 || seconds > 120) {
        hint.textContent = "请输入 5—120 之间的整数秒数。";
        return;
      }
      Office.context.document.settings.set(SETTINGS_KEY, { enabled: enabledInput.checked, seconds });
      Office.context.document.settings.saveAsync((result) => {
        hint.textContent = result.status === Office.AsyncResultStatus.Succeeded
          ? "已保存：" + (enabledInput.checked ? "每屏停留 " + seconds + " 秒。" : "自动滚动已关闭。")
          : "保存失败，请重试。";
        stop();
        done = false;
        check();
      });
    });
    document.addEventListener("answer-pending", () => {
      stop(); ready = false;
      progress.textContent = "正在准备完整解析…";
    });
    document.addEventListener("answer-failed", () => {
      stop(); progress.textContent = "讲解未完成，请重试";
    });
    document.addEventListener("answer-ready", (event) => {
      stop();
      ready = true; done = false;
      ownerSlide = event.detail.slideId;
      // app.js has replaced the previous answer, including fallback markup.
      fallback = null; fallbackHeight = 0; fallbackTop = 0;
      answer.style.overflowY = "auto";
      answer.scrollTop = 0;
      // The non-streaming response is complete. Layout is measured when the
      // dwell starts; never wait on animation frames that a host can suspend.
      progress.textContent = "完整解析已就绪";
      check();
    });
    document.addEventListener("visibilitychange", check);
    let lastHeight = answer.clientHeight;
    window.addEventListener("resize", () => {
      if (answer.clientHeight !== lastHeight) { lastHeight = answer.clientHeight; stop(); done = false; }
      check();
    });
    window.addEventListener("pagehide", stop);
    document.addEventListener("voice-state", () => { stop(); check(); });
    document.addEventListener("voice-progress", event => {
      stop(); reveal(event.detail.offset, event.detail.reset);
    });
    document.addEventListener("voice-complete", () => {
      stop();
      if (ready && visibleSlide === ownerSlide) moveTo(Math.max(0, height() - answer.clientHeight));
      done = true;
      progress.textContent = "语音及解析展示完成";
    });
    setInterval(check, 500);
    check();
  });
})();
