// Optional local speech synthesis: no new service, key, or audio upload.
(() => {
  const KEY = "ppt-ai-voice-options";
  let answer, hint, enabled, rate, owner = null, token = 0, watchdog = null, polling = false;
  let completedText = "", autoPlayed = false, previewing = false;
  const leaseKey = "ppt-ai-speech-lease", instance = Math.random().toString(36).slice(2);
  function lease(acquire) {
    try {
      const saved = JSON.parse(localStorage.getItem(leaseKey) || "null");
      if (acquire) {
        if (saved && saved.instance !== instance && saved.until > Date.now()) return false;
        localStorage.setItem(leaseKey, JSON.stringify({ instance, until: Date.now() + 3000 }));
      } else if (saved?.instance === instance) localStorage.removeItem(leaseKey);
      return true;
    } catch (_) { return false; }
  }
  function options() {
    const saved = Office.context.document.settings.get(KEY) || {};
    return { enabled: saved.enabled === true, rate: [0.85, 1, 1.15].includes(saved.rate) ? saved.rate : 1 };
  }
  function state(reading) {
    window.pptVoiceReading = reading;
    document.dispatchEvent(new Event("voice-state"));
  }
  function stop(message) {
    token++;
    clearTimeout(watchdog);
    if (window.speechSynthesis) window.speechSynthesis.cancel();
    lease(false);
    previewing = false;
    state(false);
    if (message) hint.textContent = message;
  }
  function clean(text) {
    return text.replace(/【[^】]*】/g, "").replace(/\*\*|__|`|^\s*[-#]+\s*/gm, "").replace(/GDP/g, "国内生产总值").replace(/(\d+(?:\.\d+)?)%/g, "百分之$1").trim();
  }
  function chunks(text) {
    // Short utterances are safer than one long browser speech request.
    return (text.match(/[^。！？；\n]+[。！？；\n]?/g) || [text]).flatMap(part => {
      const result = [];
      while (part.length > 120) {
        let cut = part.lastIndexOf("，", 120);
        if (cut < 40) cut = 119;
        result.push(part.slice(0, cut + 1)); part = part.slice(cut + 1);
      }
      if (part.trim()) result.push(part);
      return result;
    });
  }
  async function play(text, preview = false) {
    stop();
    if (!window.speechSynthesis || !window.SpeechSynthesisUtterance) { hint.textContent = "当前 PowerPoint 不支持本机语音，继续使用文字讲解。"; return; }
    const voices = window.speechSynthesis.getVoices();
    const voice = voices.find(v => /^zh[-_]CN/i.test(v.lang)) || voices.find(v => /^zh/i.test(v.lang));
    if (!voice) { hint.textContent = "没有检测到中文语音，请检查 Mac 系统语音；文字展示不受影响。"; return; }
    const parts = chunks(text).filter(part => clean(part));
    if (!parts.length) return;
    if (!lease(true)) { hint.textContent = "已有另一个助教窗口正在朗读，本窗口继续文字展示。"; return; }
    previewing = preview;
    const run = token;
    let index = 0;
    state(true);
    async function next() {
      if (run !== token) return;
      if (!preview) {
        const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
        if (run !== token) return;
        if (view !== "read" || slide?.id !== owner) { stop("已离开讲解页，语音停止。"); return; }
      }
      if (index >= parts.length) { stop("语音讲解完成。"); return; }
      const utterance = new window.SpeechSynthesisUtterance(clean(parts[index]));
      utterance.voice = voice; utterance.lang = voice.lang; utterance.rate = options().rate;
      utterance.onstart = () => {
        if (run !== token) return;
        clearTimeout(watchdog);
        hint.textContent = "正在朗读 " + (index + 1) + " / " + parts.length + " 段";
        // Reveal the current sentence in the complete, selectable answer.
        if (!preview) {
          const needle = parts[index].trim().replace(/[。！？；]$/, "");
          const offset = answer.textContent.indexOf(needle);
          const textNode = answer.firstChild;
          if (offset >= 0 && textNode?.nodeType === 3) {
            const range = document.createRange();
            range.setStart(textNode, offset); range.setEnd(textNode, Math.min(offset + needle.length, textNode.length));
            const rect = range.getBoundingClientRect(), viewport = answer.getBoundingClientRect();
            answer.scrollTop += rect.top - viewport.top - 8;
          }
        }
        watchdog = setTimeout(() => { if (run === token) stop("语音超时，已恢复自动文字展示。"); }, 120000);
      };
      utterance.onend = () => { if (run !== token) return; clearTimeout(watchdog); index++; next(); };
      utterance.onerror = () => { if (run === token) stop("语音播放受限或失败，已恢复自动文字展示。请课前试听。 "); };
      watchdog = setTimeout(() => { if (run === token) stop("语音未启动，已恢复自动文字展示。请在当前放映环境测试。 "); }, 7000);
      try { window.speechSynthesis.speak(utterance); }
      catch (_) { stop("语音不可用，已恢复自动文字展示。"); }
    }
    next();
  }
  Office.onReady(info => {
    if (info.host !== Office.HostType.PowerPoint) return;
    answer = document.getElementById("answer"); hint = document.getElementById("voiceStatus");
    enabled = document.getElementById("voiceEnabled"); rate = document.getElementById("voiceRate");
    enabled.checked = options().enabled; rate.value = String(options().rate);
    document.getElementById("saveVoice").addEventListener("click", () => {
      Office.context.document.settings.set(KEY, { enabled: enabled.checked, rate: Number(rate.value) });
      Office.context.document.settings.saveAsync(result => {
        stop(result.status === Office.AsyncResultStatus.Succeeded ? "语音设置已保存。请课前试听。" : "语音设置保存失败，请重试。");
      });
    });
    document.getElementById("testVoice").addEventListener("click", () => play("你好，我是宏观经济学 AI 助教。请确认教室音响能够听到这段声音。", true));
    document.getElementById("stopVoice").addEventListener("click", () => stop("语音已停止。"));
    document.addEventListener("answer-pending", () => { stop(); completedText = ""; autoPlayed = false; });
    document.addEventListener("answer-ready", event => { owner = event.detail.slideId; completedText = answer.textContent; autoPlayed = false; });
    // Only the slideshow instance speaks. Never treat document.hidden as exit.
    setInterval(async () => {
      if (polling) return;
      polling = true;
      try {
        const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
        if (window.pptVoiceReading && !lease(true)) { stop("另一助教窗口已接管语音。"); return; }
        if (previewing) return;
        if (view !== "read" || slide?.id !== owner) { if (window.pptVoiceReading) stop("已离开放映讲解页，语音停止。"); return; }
        if (completedText && options().enabled && !autoPlayed) { autoPlayed = true; play(completedText); }
      } finally { polling = false; }
    }, 500);
    window.addEventListener("pagehide", () => stop());
  });
})();
