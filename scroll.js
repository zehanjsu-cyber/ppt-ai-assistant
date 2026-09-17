// Automatic answer scrolling stays inside the add-in, not the PowerPoint slide.
(() => {
  const SETTINGS_KEY = "ppt-ai-scroll-options";
  let answer, enabledInput, secondsInput, hint;
  let ownerSlide = null, visibleSlide = null, timer = null;
  let checking = false, generation = 0, ready = false, done = false, phase = "idle";
  let progress;

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
    if (phase !== "idle" || !ready || done || document.hidden || !options().enabled || !visibleSlide || visibleSlide !== ownerSlide || answer.clientHeight <= 0) return;
    const atEnd = answer.scrollHeight <= answer.clientHeight + answer.scrollTop + 2;
    progress.textContent = atEnd ? "最后一屏 · 阅读后停止" : "完整解析 · 每屏停留 " + options().seconds + " 秒";
    const token = generation;
    phase = "dwell";
    timer = setTimeout(async () => {
      phase = "checking";
      const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
      if (token !== generation) return;
      if (document.hidden || view !== "read" || !slide || slide.id !== ownerSlide || !options().enabled) { stop(); return; }
      if (answer.scrollHeight <= answer.clientHeight + answer.scrollTop + 2) {
        done = true; timer = null; phase = "idle";
        progress.textContent = "全部解析已展示";
        return;
      }
      phase = "moving";
      const overlap = Math.min(60, answer.clientHeight * 0.15);
      answer.scrollTo({
        top: Math.min(answer.scrollHeight - answer.clientHeight,
          answer.scrollTop + Math.max(1, answer.clientHeight - overlap)),
        behavior: "smooth",
      });
      // Start a fresh dwell only after movement has settled, not after a fixed delay.
      let previous = -1, stable = 0;
      function settle() {
        if (token !== generation) return;
        stable = Math.abs(answer.scrollTop - previous) < .5 ? stable + 1 : 0;
        previous = answer.scrollTop;
        if (stable >= 4) { timer = null; phase = "idle"; schedule(); }
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
      const current = view === "read" && slide && !document.hidden ? slide.id : null;
      document.body.classList.toggle("show-view", view === "read");
      if (current !== visibleSlide) {
        stop();
        visibleSlide = current;
        answer.scrollTop = 0;
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
      ready = false; done = false;
      ownerSlide = event.detail.slideId;
      answer.scrollTop = 0;
      const token = generation;
      requestAnimationFrame(() => requestAnimationFrame(() => {
        if (token !== generation) return;
        ready = true;
        progress.textContent = "完整解析已就绪";
        check();
      }));
    });
    document.addEventListener("visibilitychange", () => { stop(); check(); });
    window.addEventListener("resize", () => { stop(); done = false; check(); });
    window.addEventListener("pagehide", stop);
    setInterval(check, 500);
    check();
  });
})();
