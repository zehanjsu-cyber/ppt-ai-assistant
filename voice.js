// Optional local speech synthesis: no new service, key, or audio upload.
(() => {
  const KEY = "ppt-ai-voice-options";
  let answer, hint, enabled, rate, owner = null, token = 0, watchdog = null, polling = false;
  let completedText = "", autoPlayed = false, previewing = false, visibleOwner = false;
  let voiceSelect, repeatSelect, replayButton;
  let readMisses = 0, retryAt = 0;
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
    return { enabled: saved.enabled === true, rate: [0.85, 1, 1.15].includes(saved.rate) ? saved.rate : 1,
      voice: typeof saved.voice === "string" ? saved.voice : "", repeat: [1, 2, 0].includes(saved.repeat) ? saved.repeat : 1 };
  }
  function state(reading) {
    window.pptVoiceReading = reading;
    document.dispatchEvent(new Event("voice-state"));
  }
  function stop(message) {
    token++;
    clearTimeout(watchdog);
    if (window.pptVoiceReading && window.speechSynthesis) window.speechSynthesis.cancel();
    lease(false);
    previewing = false;
    state(false);
    if (message) hint.textContent = message;
  }
  function clean(text) {
    return text
      .replace(/【[^】]*】|\[\d+(?:[†,，\-]\S*?)?\]/g, "")
      .replace(/!\[[^\]]*\]\([^)]*\)/g, "")
      .replace(/\[([^\]]+)\]\([^)]*\)/g, "$1")
      .replace(/https?:\/\/\S+/g, "")
      // Decorative emoji only: retain mathematical signs such as +, −, =, >.
      .replace(/[\p{Extended_Pictographic}\p{Regional_Indicator}\uFE0F\u200D\u20E3\u2610-\u2612\u2713-\u2718\u2022\u25CF\u25CB]/gu, "")
      .replace(/\*\*|__|`|^\s*[-#>]+\s*|^\s*\d+[.)、]\s*/gm, "")
      .replace(/\|/g, "，").replace(/GDP/gi, "国内生产总值")
      .replace(/(\d+(?:\.\d+)?)%/g, "百分之$1")
      .replace(/\s+/g, " ").trim();
  }
  function voiceId(voice) { return voice.voiceURI || voice.name + "|" + voice.lang; }
  function availableVoices() {
    return (window.speechSynthesis?.getVoices() || []).filter(v => /^zh/i.test(v.lang) && v.localService !== false);
  }
  function refreshVoices() {
    const chosen = voiceSelect.value || options().voice;
    voiceSelect.replaceChildren();
    const automatic = document.createElement("option");
    automatic.value = ""; automatic.textContent = "自动选择本机普通话声音";
    voiceSelect.appendChild(automatic);
    for (const voice of availableVoices()) {
      const option = document.createElement("option");
      option.value = voiceId(voice); option.textContent = voice.name + "（" + voice.lang + "）";
      voiceSelect.appendChild(option);
    }
    voiceSelect.value = availableVoices().some(v => voiceId(v) === chosen) ? chosen : "";
  }
  function chunks(text) {
    // Short utterances are safer than one long browser speech request.
    const units = (text.match(/[^。！？；\n]+[。！？；\n]*/g) || [text]).flatMap(part => {
      const result = [];
      while (part.length > 200) {
        let cut = part.lastIndexOf("，", 180);
        if (cut < 60) cut = 179;
        result.push(part.slice(0, cut + 1)); part = part.slice(cut + 1);
      }
      if (part.trim()) result.push(part);
      return result;
    });
    const result = []; let current = "";
    for (const unit of units) {
      if (!clean(unit)) { current += unit; continue; }
      if (current && current.length + unit.length > 200) { result.push(current); current = ""; }
      current += unit;
      if (/\n\s*\n$/.test(unit) && current.length > 50) { result.push(current); current = ""; }
    }
    if (current.trim()) result.push(current);
    return result;
  }
  async function play(text, preview = false, direct = false) {
    if (!window.speechSynthesis || !window.SpeechSynthesisUtterance) { hint.textContent = "当前 PowerPoint 不支持本机语音，继续使用文字讲解。"; return true; }
    const voices = availableVoices();
    const selection = preview ? voiceSelect.value : options().voice;
    const voice = voices.find(v => voiceId(v) === selection) || voices.find(v => /^zh[-_]CN/i.test(v.lang) && v.default) || voices.find(v => /^zh[-_]CN/i.test(v.lang)) || voices[0];
    if (!voice) { hint.textContent = "正在等待本机中文声音加载；也可点击“播放 / 重听”重试。"; retryAt = Date.now() + 1000; return false; }
    const parts = chunks(text).filter(part => clean(part));
    let searchFrom = 0;
    const offsets = parts.map(part => { const at = text.indexOf(part, searchFrom); searchFrom = Math.max(searchFrom, at + part.length); return Math.max(0, at); });
    if (!parts.length) return true;
    if (!lease(true)) { hint.textContent = "已有另一个助教窗口正在朗读，本窗口继续文字展示。"; retryAt = Date.now() + 3000; return false; }
    stop(); lease(true);
    previewing = preview;
    const run = token;
    let index = 0, round = 1, first = true;
    state(true);
    async function next() {
      if (run !== token) return;
      if (!preview && !(direct && first)) {
        const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
        if (run !== token) return;
        if (!view || (view === "read" && !slide)) {
          watchdog = setTimeout(() => { if (run === token) next(); }, 500);
          if (++readMisses >= 4) stop("暂时无法确认放映页，语音暂停；请点“播放 / 重听”重试。");
          return;
        }
        readMisses = 0;
        if (view !== "read" || slide?.id !== owner) { stop("已离开讲解页，语音停止。"); return; }
      }
      first = false;
      if (index >= parts.length) {
        if (!preview && (options().repeat === 0 || round < options().repeat)) {
          round++; index = 0;
          hint.textContent = "本遍已完成，5 秒后重听；翻页可停止。";
          watchdog = setTimeout(() => { if (run === token) { answer.scrollTop = 0; next(); } }, 5000);
          return;
        }
        if (!preview) document.dispatchEvent(new Event("voice-complete"));
        stop("语音讲解完成。可点“重听”，或翻走后返回自动重听。"); return;
      }
      const utterance = new window.SpeechSynthesisUtterance(clean(parts[index]));
      utterance.voice = voice; utterance.lang = voice.lang;
      utterance.rate = preview ? Number(rate.value) : options().rate;
      utterance.onstart = () => {
        if (run !== token) return;
        clearTimeout(watchdog);
        hint.textContent = "正在朗读 " + (index + 1) + " / " + parts.length + " 段";
        if (!preview) document.dispatchEvent(new CustomEvent("voice-progress", { detail: { offset: offsets[index], reset: index === 0 } }));
        watchdog = setTimeout(() => { if (run === token) stop("语音超时，已恢复自动文字展示。"); }, 120000);
      };
      utterance.onboundary = event => {
        if (run !== token || preview || !Number.isInteger(event.charIndex)) return;
        // Translate cleaned speech positions back to the original displayed text.
        const raw = parts[index];
        let low = 0, high = raw.length;
        while (low < high) { const mid = (low + high) >> 1; if (clean(raw.slice(0, mid)).length < event.charIndex) low = mid + 1; else high = mid; }
        document.dispatchEvent(new CustomEvent("voice-progress", { detail: { offset: offsets[index] + low, reset: false } }));
      };
      utterance.onend = () => {
        if (run !== token) return;
        clearTimeout(watchdog); index++;
        // Natural paragraph breathing, rather than restarting every short line.
        watchdog = setTimeout(() => { if (run === token) next(); }, 300);
      };
      utterance.onerror = () => { if (run === token) stop("语音播放受限或失败，已恢复自动文字展示。请课前试听。 "); };
      watchdog = setTimeout(() => { if (run === token) stop("语音未启动，已恢复自动文字展示。请在当前放映环境测试。 "); }, 7000);
      try { window.speechSynthesis.resume?.(); window.speechSynthesis.speak(utterance); }
      catch (_) { stop("语音不可用，已恢复自动文字展示。"); }
    }
    next();
    return true;
  }
  Office.onReady(info => {
    if (info.host !== Office.HostType.PowerPoint) return;
    answer = document.getElementById("answer");
    const settingsHint = document.getElementById("voiceStatus"), liveHint = document.getElementById("voiceLiveStatus");
    hint = { get textContent() { return settingsHint.textContent; }, set textContent(value) { settingsHint.textContent = value; if (liveHint) liveHint.textContent = value; } };
    enabled = document.getElementById("voiceEnabled"); rate = document.getElementById("voiceRate");
    voiceSelect = document.getElementById("voiceSelect"); repeatSelect = document.getElementById("voiceRepeat");
    replayButton = document.getElementById("replayVoice");
    enabled.checked = options().enabled; rate.value = String(options().rate);
    repeatSelect.value = String(options().repeat); refreshVoices();
    window.speechSynthesis?.addEventListener?.("voiceschanged", refreshVoices);
    function saveOptions() {
      Office.context.document.settings.set(KEY, { enabled: enabled.checked, rate: Number(rate.value), voice: voiceSelect.value, repeat: Number(repeatSelect.value) });
      Office.context.document.settings.saveAsync(result => {
        stop(result.status === Office.AsyncResultStatus.Succeeded ? (enabled.checked ? "语音已开启：放映中自动朗读；无声音时点“播放 / 重听”。" : "语音已关闭。") : "语音设置保存失败，请重试。");
        autoPlayed = false; retryAt = 0;
      });
    }
    document.getElementById("saveVoice").addEventListener("click", saveOptions);
    enabled.addEventListener("change", saveOptions);
    document.getElementById("testVoice").addEventListener("click", () => play("你好，我是宏观经济学 AI 助教。请确认教室音响能够听到这段声音。", true, true));
    document.getElementById("stopVoice").addEventListener("click", () => stop("语音已停止。"));
    replayButton.addEventListener("click", () => {
      if (!completedText) return;
      // Submit speech inside the click handler, before any await loses activation.
      autoPlayed = true; answer.scrollTop = 0; play(completedText, !visibleOwner, true);
    });
    document.addEventListener("answer-pending", () => { stop(); completedText = ""; autoPlayed = false; replayButton.disabled = true; });
    document.addEventListener("answer-ready", event => { owner = event.detail.slideId; completedText = answer.textContent; autoPlayed = false; replayButton.disabled = false; });
    // Only the slideshow instance speaks. Never treat document.hidden as exit.
    setInterval(async () => {
      if (polling) return;
      polling = true;
      try {
        const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
        if (window.pptVoiceReading && !lease(true)) { stop("另一助教窗口已接管语音。"); return; }
        if (previewing) return;
        if (!view || (view === "read" && !slide)) {
          if (++readMisses >= 4 && window.pptVoiceReading) stop("暂时无法确认放映页，语音暂停；请点“播放 / 重听”重试。");
          return;
        }
        readMisses = 0;
        const onOwner = view === "read" && slide?.id === owner;
        if (!onOwner) {
          if (visibleOwner) autoPlayed = false;
          visibleOwner = false;
          if (window.pptVoiceReading) stop("已离开放映讲解页，语音停止。"); return;
        }
        visibleOwner = true;
        if (completedText && options().enabled && !autoPlayed && Date.now() >= retryAt) { autoPlayed = await play(completedText) === true; }
      } finally { polling = false; }
    }, 500);
    window.addEventListener("pagehide", () => stop());
  });
})();
