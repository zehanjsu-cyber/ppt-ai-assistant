// Automatic answer scrolling stays inside the add-in, not the PowerPoint slide.
(() => {
  const SETTINGS_KEY = "ppt-ai-scroll-options";
  let answer, enabledInput, secondsInput, hint;
  let ownerSlide = null, visibleSlide = null, timer = null;
  let checking = false, generation = 0;

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
    if (answer) answer.scrollTo({ top: answer.scrollTop, behavior: "instant" });
  }

  function schedule() {
    if (timer || !options().enabled || !visibleSlide || visibleSlide !== ownerSlide) return;
    if (answer.scrollHeight <= answer.clientHeight + answer.scrollTop + 2) return;
    const token = generation;
    timer = setTimeout(async () => {
      timer = null;
      const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
      if (token !== generation || view !== "read" || !slide || slide.id !== ownerSlide || !options().enabled) return;
      const overlap = Math.min(60, answer.clientHeight * 0.15);
      answer.scrollTo({
        top: Math.min(answer.scrollHeight - answer.clientHeight,
          answer.scrollTop + Math.max(1, answer.clientHeight - overlap)),
        behavior: "smooth",
      });
      // Allow the smooth movement to finish before starting the next dwell.
      timer = setTimeout(() => {
        timer = null;
        if (token === generation) schedule();
      }, 1000);
    }, options().seconds * 1000);
  }

  async function check() {
    if (checking) return;
    checking = true;
    try {
      const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
      const current = view === "read" && slide && !document.hidden ? slide.id : null;
      document.body.classList.toggle("show-view", view === "read");
      if (current !== visibleSlide) {
        stop();
        visibleSlide = current;
        answer.scrollTop = 0;
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
        check();
      });
    });
    new MutationObserver(async () => {
      stop();
      ownerSlide = null;
      answer.scrollTop = 0;
      const token = generation;
      const slide = await getCurrentSlideInfo();
      if (token !== generation) return;
      ownerSlide = slide ? slide.id : null;
      check();
    }).observe(answer, { childList: true, subtree: true, characterData: true });
    document.addEventListener("visibilitychange", () => { stop(); check(); });
    window.addEventListener("resize", () => { stop(); check(); });
    window.addEventListener("pagehide", stop);
    setInterval(check, 500);
    check();
  });
})();
