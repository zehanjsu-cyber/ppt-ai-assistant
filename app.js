const API_URL = "https://yuanqi.tencent.com/openapi/v1/agent/chat/completions";
const ASSISTANT_ID = "2096407988983757888";
const STORAGE_KEY = "yuanqi-app-key";
const AUTO_SLIDES_KEY = "ppt-ai-auto-slide-ids";
const RAIN_STATS_ENABLED_KEY = "ppt-ai-rain-stats-enabled";
const RAIN_STATS_URL = "https://127.0.0.1:19789/latest";
const RAIN_HEALTH_URL = "https://127.0.0.1:19789/health";
const RAIN_STATS_MAX_AGE_MS = 10 * 60 * 1000;
const RAIN_STATS_TIMEOUT_MS = 5000;

const els = {};
let conversation = [];
let isBusy = false;
let wasReadView = false;
let answerSlideId = null;
const autoTriggeredSlides = new Set();

Office.onReady((info) => {
  Object.assign(els, {
    explain: document.getElementById("explainButton"),
    status: document.getElementById("status"),
    answer: document.getElementById("answer"),
    settings: document.getElementById("settings"),
    settingsButton: document.getElementById("settingsButton"),
    appKey: document.getElementById("appKey"),
    saveKey: document.getElementById("saveKey"),
    clearKey: document.getElementById("clearKey"),
    toggleAuto: document.getElementById("toggleAuto"),
    autoStatus: document.getElementById("autoStatus"),
    rainStatsEnabled: document.getElementById("rainStatsEnabled"),
    rainStatsStatus: document.getElementById("rainStatsStatus"),
    rainStatsLiveStatus: document.getElementById("rainStatsLiveStatus"),
    checkRainStats: document.getElementById("checkRainStats"),
    followupInput: document.getElementById("followupInput"),
    followupButton: document.getElementById("followupButton"),
    followupQuestion: document.getElementById("followupQuestion"),
  });

  els.settingsButton.addEventListener("click", () => els.settings.classList.toggle("hidden"));
  els.saveKey.addEventListener("click", saveKey);
  els.clearKey.addEventListener("click", clearKey);
  els.toggleAuto.addEventListener("click", toggleCurrentSlideAuto);
  els.rainStatsEnabled.checked = localStorage.getItem(RAIN_STATS_ENABLED_KEY) === "1";
  els.rainStatsEnabled.addEventListener("change", saveRainStatsSetting);
  els.checkRainStats.addEventListener("click", checkRainStatsConnection);
  updateRainStatsLive(null, els.rainStatsEnabled.checked ? "尚未读取作答分布" : "功能未启用");
  els.explain.addEventListener("click", explainCurrentSlide);
  els.followupButton.addEventListener("click", askFollowup);
  els.followupInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.isComposing && event.keyCode !== 229) {
      event.preventDefault(); askFollowup();
    }
  });

  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("请在 PowerPoint 中打开此加载项。", true);
    return;
  }

  els.explain.disabled = false;
  setStatus("已就绪：切到题目页后点击一次按钮。", false);
  if (!localStorage.getItem(STORAGE_KEY)) els.settings.classList.remove("hidden");
  refreshAutoControl();
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    refreshAutoControl,
    () => {}
  );
  setInterval(checkAutoExplain, 800);
});

function setStatus(message, isError = false) {
  els.status.textContent = message;
  els.status.classList.toggle("error", isError);
}

function saveRainStatsSetting() {
  localStorage.setItem(RAIN_STATS_ENABLED_KEY, els.rainStatsEnabled.checked ? "1" : "0");
  els.rainStatsStatus.textContent = els.rainStatsEnabled.checked
    ? "已开启。讲解时会优先使用最近一次有效作答分布。"
    : "已关闭。AI 仅根据当前 PPT 题目讲解。";
  updateRainStatsLive(null, els.rainStatsEnabled.checked ? "已开启，等待作答分布" : "功能未启用");
}

function updateRainStatsLive(stats, message) {
  if (!els.rainStatsLiveStatus) return;
  els.rainStatsLiveStatus.classList.toggle("used", Boolean(stats));
  els.rainStatsLiveStatus.classList.toggle("warning", !stats && els.rainStatsEnabled?.checked);
  els.rainStatsLiveStatus.textContent = stats
    ? `学情数据：已使用（${stats.options.map(item => `${item.label} ${item.count ?? "?"}人`).join("，")}）`
    : `学情数据：未使用（${message}）`;
}

function normalizeRainStats(payload) {
  if (!payload || !Array.isArray(payload.options) || payload.options.length < 2) return null;
  const capturedAt = Date.parse(payload.captured_at || "");
  if (!Number.isFinite(capturedAt) || capturedAt > Date.now() + 30000 || Date.now() - capturedAt > RAIN_STATS_MAX_AGE_MS) return null;
  const options = payload.options
    .map((item) => ({
      label: String(item.label || "").trim().toUpperCase(),
      count: Number.isInteger(item.count) && item.count >= 0 ? item.count : null,
      percent: typeof item.percent === "number" && item.percent >= 0 && item.percent <= 100 ? item.percent : null,
      correct: item.correct === true,
    }))
    .filter((item) => /^[A-Z]$/.test(item.label));
  const submitted = Number.isInteger(payload.submitted) && payload.submitted >= 0 ? payload.submitted : null;
  if (options.length < 2 || options.some(item => item.count === null) || new Set(options.map(item => item.label)).size !== options.length) return null;
  if (submitted === 0 || options.every(item => item.count === 0)) return null;
  if (submitted !== null && options.some(item => item.count > submitted)) return null;
  return {
    capturedAt: new Date(capturedAt),
    submitted,
    options,
    questionText: typeof payload.question_text === "string" ? payload.question_text : "",
  };
}

function normalizedQuestion(text) {
  return String(text || "").normalize("NFKC").toUpperCase()
    .replace(/学情数据[^\n]*/g, "")
    .replace(/[^\p{Script=Han}A-Z0-9]/gu, "");
}

function questionMatches(statsText, slideText) {
  const source = normalizedQuestion(statsText);
  const target = normalizedQuestion(slideText);
  if (source.length < 12 || target.length < 12) return false;
  const units = new Set();
  for (let index = 0; index < target.length - 3; index += 2) units.add(target.slice(index, index + 4));
  let matches = 0;
  for (const unit of units) if (source.includes(unit)) matches++;
  if (units.size < 3 || matches < 3) return false;
  if (matches / units.size >= 0.35) return true;
  // A Terminal window can cover part of the answer options in the screenshot.
  // Allow that narrow case only when the full, non-generic question stem is
  // visible and the remaining option text still contributes overlap.
  const stem = normalizedQuestion(String(slideText).split(/[?？]/, 1)[0]);
  return stem.length >= 18 && source.includes(stem) && matches / units.size >= 0.30;
}

function rainConnectionError(error) {
  if (error?.name === "AbortError") return "识别助手连接超时（5秒）";
  const detail = String(error?.message || error || "未知错误").slice(0, 100);
  return `PowerPoint 未能读取识别助手返回内容（${detail}）`;
}

function rainEmptyMessage(health) {
  if (health?.last_capture_error?.includes("屏幕录制失败")) return "识别助手没有屏幕录制权限；请在系统设置中允许后重启助手";
  if (health?.last_capture_error?.includes("已看到非零作答窗口")) return health.last_capture_error;
  if (health?.last_capture_error && !health.last_capture_error.includes("未发现") && !health.last_capture_error.includes("尚未识别")) {
    return `识别助手截图失败：${health.last_capture_error}`;
  }
  return "未捕获非零作答分布；请在助手运行时打开雨课堂“作答情况”窗口，等终端显示“已记录作答分布”后再翻到讲解页";
}

async function getRainStatsResult(questionTexts = []) {
  if (!els.rainStatsEnabled?.checked) return { stats: null, reason: "功能未启用" };
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), RAIN_STATS_TIMEOUT_MS);
  try {
    const response = await fetch(RAIN_STATS_URL, { cache: "no-store", signal: controller.signal });
    if (!response.ok) return { stats: null, reason: `识别助手返回 ${response.status}` };
    const payload = await response.json();
    const stats = normalizeRainStats(payload);
    if (stats && questionTexts.length && !questionTexts.some(text => questionMatches(stats.questionText, text))) {
      return { stats: null, reason: "统计题目与本题不匹配，已防止串题" };
    }
    if (stats) return { stats, reason: "" };
    if (!payload?.captured_at) {
      try {
        const healthResponse = await fetch(RAIN_HEALTH_URL, { cache: "no-store", signal: controller.signal });
        if (healthResponse.ok) return { stats: null, reason: rainEmptyMessage(await healthResponse.json()) };
      } catch (_) { /* The /latest response already confirms connectivity. */ }
      return { stats: null, reason: rainEmptyMessage(null) };
    }
    return { stats: null, reason: "最近数据不完整或已经过期" };
  } catch (error) {
    return { stats: null, reason: rainConnectionError(error) };
  } finally {
    clearTimeout(timeout);
  }
}

function parseSlideQuestion(rawText) {
  const answers = [];
  const followups = [];
  const lines = [];
  for (const rawLine of String(rawText || "").split(/\r?\n/)) {
    const line = rawLine.normalize("NFKC").trim();
    if (/^标准答案\s*:/i.test(line)) {
      const match = /^标准答案\s*:\s*([A-D])\s*$/i.exec(line);
      if (!match) throw new Error("标准答案格式应为独立一行：标准答案：B（仅 A、B、C、D 单选）。");
      answers.push(match[1].toUpperCase());
    } else if (/^课堂追问\s*:/i.test(line)) {
      const match = /^课堂追问\s*:\s*(.+)$/i.exec(line);
      if (!match) throw new Error("课堂追问格式应为独立一行：课堂追问：为什么 C 不对？");
      followups.push(match[1].trim());
    } else {
      lines.push(rawLine);
    }
  }
  if (new Set(answers).size > 1) throw new Error("本页有多个不同的标准答案，请只保留一个。");
  if (followups.length > 1) throw new Error("本页只保留一个课堂追问；需要更多追问请另用一页。");
  if (followups.length && !answers.length) throw new Error("追问页缺少“标准答案：B”，已停止自动追问。");
  return { question: lines.join("\n").trim(), correctAnswer: answers[0] || null, followup: followups[0] || null };
}

function formatRainStats(stats, correctAnswer = null) {
  const rows = stats.options.map((item) => {
    const details = [];
    if (item.count !== null) details.push(`${item.count}人`);
    if (item.percent !== null) details.push(`${item.percent}%`);
    if (correctAnswer === item.label) details.push("本页标注的正确选项");
    return `${item.label}：${details.join("，") || "已识别"}`;
  });
  if (stats.submitted !== null) rows.unshift(`提交人数：${stats.submitted}人`);
  return rows.join("\n");
}

function buildRainTeachingGuidance(stats, correctAnswer) {
  const wrongChoices = stats.options
    .filter((item) => item.label !== correctAnswer && item.count > 0)
    .sort((left, right) => right.count - left.count)
    .map((item) => `${item.label}（${item.count}人）`);
  if (!wrongChoices.length) {
    return "本题没有人选错。不要编造错因；在正常讲解后用一句话提醒判断关键即可。";
  }
  return `本班有学生选了错误选项：${wrongChoices.join("、")}。正常讲完正确答案和核心原理后，必须单独加一段“本班易错点”（最多两句）：先准确点名这些选项及人数，再依据课程知识指出一种可能的概念混淆，并给出一个具体的判断提醒。不要仅重复逐项判断，不要推断学生个人真实想法；1人作答不得说“多数人”或“普遍”。`;
}

function parseIndependentAnswer(text) {
  const value = String(text || "").normalize("NFKC").trim();
  const match = /^(?:独立判断|答案)\s*[:：]\s*([A-D]|无法判断)\s*[。.]?$/i.exec(value);
  return match && match[1] !== "无法判断" ? match[1].toUpperCase() : null;
}

async function verifyTeacherAnswer(key, question, correctAnswer) {
  if (!correctAnswer) return;
  const checkPrompt = `请只依据你已配置的课程知识库，独立判断下面这道单选题的正确选项。此阶段不要参考教师答案或学生作答人数，也不要展开讲解。只输出一行“独立判断：A”（A、B、C、D之一）；题意不清、知识不足或有多个合理选项时只输出“独立判断：无法判断”。\n\n题目：\n${question}`;
  const result = await callYuanqi(key, [{ role: "user", content: [{ type: "text", text: checkPrompt }] }]);
  const independent = parseIndependentAnswer(result);
  if (!independent) throw new Error("AI 无法明确独立判断本题答案，已停止讲解。请核对题目和标准答案。");
  if (independent !== correctAnswer) throw new Error(`答案冲突：PPT 标准答案为 ${correctAnswer}，AI 独立判断为 ${independent}。已停止讲解，请教师核对。`);
}

async function checkRainStatsConnection() {
  const hasKey = Boolean(localStorage.getItem(STORAGE_KEY));
  let health = null;
  let healthError = "";
  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), RAIN_STATS_TIMEOUT_MS);
    try {
      const response = await fetch(RAIN_HEALTH_URL, { cache: "no-store", signal: controller.signal });
      if (response.ok) health = await response.json();
    } finally { clearTimeout(timeout); }
  } catch (error) { healthError = rainConnectionError(error); }
  const result = await getRainStatsResult();
  const stats = result.stats;
  const checks = [
    hasKey ? "AI密钥已保存" : "AI密钥未保存",
    els.rainStatsEnabled.checked ? "学情开关已开" : "学情开关已关",
    health?.ok ? "识别助手已连接" : `识别助手未连接：${healthError || "返回内容异常"}`,
    stats ? "已捕获分布（尚未与题目核对）" : result.reason,
    window.speechSynthesis ? "本机语音接口可用（需课前试听）" : "本机语音接口不可用",
  ];
  els.rainStatsStatus.textContent = checks.join("；") + "。";
  updateRainStatsLive(null, stats ? "已捕获分布，讲解时再核对题目" : result.reason);
}

function saveKey() {
  const value = els.appKey.value.trim();
  if (!value) return setStatus("请先粘贴 AppKey。", true);
  localStorage.setItem(STORAGE_KEY, value);
  els.appKey.value = "";
  els.settings.classList.add("hidden");
  setStatus("AppKey 已保存在本机。", false);
}

function clearKey() {
  localStorage.removeItem(STORAGE_KEY);
  els.appKey.value = "";
  setStatus("AppKey 已清除。", false);
}

function getCurrentSlideInfo() {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.SlideRange,
      (result) => {
        if (
          result.status === Office.AsyncResultStatus.Succeeded &&
          result.value?.slides?.length
        ) {
          const slide = result.value.slides[0];
          resolve({ id: String(slide.id), index: slide.index - 1 });
        } else {
          resolve(null);
        }
      }
    );
  });
}

function getActiveView() {
  return new Promise((resolve) => {
    Office.context.document.getActiveViewAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : null);
    });
  });
}

function getAutoSlideIds() {
  const value = Office.context.document.settings.get(AUTO_SLIDES_KEY);
  return Array.isArray(value) ? value.map(String) : [];
}

function saveAutoSlideIds(ids) {
  return new Promise((resolve, reject) => {
    Office.context.document.settings.set(AUTO_SLIDES_KEY, ids);
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error.message));
    });
  });
}

async function refreshAutoControl() {
  const [view, slide] = await Promise.all([getActiveView(), getCurrentSlideInfo()]);
  if (view === "read") {
    els.toggleAuto.disabled = true;
    els.autoStatus.textContent = "放映中；请在编辑模式设置自动讲解页。";
    return;
  }
  els.toggleAuto.disabled = !slide;
  const enabled = slide && getAutoSlideIds().includes(slide.id);
  els.toggleAuto.textContent = enabled ? "取消本页自动讲解" : "将本页设为自动讲解页";
  els.autoStatus.textContent = enabled
    ? "本页已开启：放映翻到这里会自动讲解一次。"
    : "备课时在需要自动讲解的页面开启。";
}

async function toggleCurrentSlideAuto() {
  const slide = await getCurrentSlideInfo();
  if (!slide) return setStatus("请先进入需要自动讲解的幻灯片。", true);
  const ids = getAutoSlideIds();
  const next = ids.includes(slide.id)
    ? ids.filter((id) => id !== slide.id)
    : [...ids, slide.id];
  try {
    await saveAutoSlideIds(next);
    await refreshAutoControl();
    setStatus(next.includes(slide.id) ? "本页自动讲解已开启。" : "本页自动讲解已取消。", false);
  } catch (error) {
    setStatus(error.message || String(error), true);
  }
}

async function checkAutoExplain() {
  if (isBusy) return;
  const view = await getActiveView();
  if (view !== "read") {
    if (wasReadView) autoTriggeredSlides.clear();
    wasReadView = false;
    return;
  }
  if (!wasReadView) autoTriggeredSlides.clear();
  wasReadView = true;
  const slide = await getCurrentSlideInfo();
  if (!slide || !getAutoSlideIds().includes(slide.id) || autoTriggeredSlides.has(slide.id)) return;
  autoTriggeredSlides.add(slide.id);
  setStatus("已进入自动讲解页，正在开始讲解…", false);
  await explainCurrentSlide();
}

async function readCurrentSlideText(currentSlide) {
  return PowerPoint.run(async (context) => {
    let slide;
    if (currentSlide !== null) {
      slide = context.presentation.slides.getItemAt(currentSlide.index);
    } else {
      const selectedSlides = context.presentation.getSelectedSlides();
      const slideCount = selectedSlides.getCount();
      selectedSlides.load("items");
      await context.sync();
      if (slideCount.value < 1) throw new Error("没有识别到正在显示的幻灯片。");
      slide = selectedSlides.items[0];
    }

    const shapes = slide.shapes;
    shapes.load("items");
    await context.sync();

    const textShapes = shapes.items.filter((shape) =>
      shape.type === PowerPoint.ShapeType.textBox ||
      shape.type === PowerPoint.ShapeType.geometricShape ||
      shape.type === PowerPoint.ShapeType.placeholder
    );
    textShapes.forEach((shape) => shape.textFrame.load("hasText,textRange/text"));
    await context.sync();

    const lines = [];
    for (const shape of textShapes) {
      if (shape.textFrame.hasText) {
        const text = shape.textFrame.textRange.text.trim();
        if (text) lines.push(text);
      }
    }
    return lines.join("\n").trim();
  });
}

async function explainCurrentSlide() {
  if (isBusy) return;
  const key = localStorage.getItem(STORAGE_KEY);
  if (!key) {
    els.settings.classList.remove("hidden");
    return setStatus("首次使用：请粘贴并保存 AppKey。", true);
  }

  isBusy = true;
  document.dispatchEvent(new Event("answer-pending"));
  conversation = [];
  answerSlideId = null;
  els.answer.textContent = "正在核对题目…";
  setBusy(true, "正在读取当前页…");
  try {
    const sourceSlide = await getCurrentSlideInfo();
    if (!sourceSlide) throw new Error("没有识别到正在显示的幻灯片，请重试。");
    const slideText = await readCurrentSlideText(sourceSlide);
    const { question, correctAnswer, followup } = parseSlideQuestion(slideText);
    if (!question) throw new Error("当前页没有读取到文字。图片题请把题目文字放在一个文本框中。");

    if (correctAnswer) {
      setBusy(true, "正在独立核对标准答案…");
      await verifyTeacherAnswer(key, question, correctAnswer);
    }

    const rainResult = await getRainStatsResult([question]);
    if (rainResult.stats && !correctAnswer) {
      rainResult.stats = null;
      rainResult.reason = "本页未标注教师标准答案，已避免把 AI 自判当作学情判错";
    }
    if (rainResult.stats && correctAnswer && !rainResult.stats.options.some(item => item.label === correctAnswer)) {
      rainResult.stats = null;
      rainResult.reason = `统计中缺少标准答案 ${correctAnswer} 选项，已防止误判`;
    }
    const rainStats = rainResult.stats;
    updateRainStatsLive(rainStats, rainResult.reason);
    const answerContext = correctAnswer
      ? `\n\n教师在本页标注的标准答案是 ${correctAnswer}。这是判定对错的依据，不要自行改判；如果题目与该答案明显矛盾，请指出矛盾，不要编造理由。`
      : "\n\n本页未标注标准答案，请独立判断；不要仅凭作答人数推断正确选项。";
    const statsContext = rainStats
      ? `\n\n这是刚才雨课堂的全班汇总数据：\n${formatRainStats(rainStats, correctAnswer)}\n\n${buildRainTeachingGuidance(rainStats, correctAnswer)}只把分布表述为可能的误区，不提及任何学生个人。`
      : "";
    const task = followup
      ? `这是教师预设的课堂追问：${followup}\n\n请直接回答这个追问，不要重新完整讲一遍原题。若追问的前提与教师标注的标准答案冲突，先指出冲突，不得编造支持理由。`
      : "请先明确给出正确答案，再解释核心原理；有选项时逐项判断，最后用一句话总结考点。";
    const prompt = `你是大学宏观经济学课堂的AI课程助教。原题如下：\n\n${question}${answerContext}${statsContext}\n\n${task}\n\n要求：优先依据你已配置的课程知识库解释原理；知识库没有支持的细节不要编造；不虚构题目中没有的数据；控制在课堂60—90秒可讲完。`;
    conversation = [{ role: "user", content: [{ type: "text", text: prompt }] }];
    const answer = await callYuanqi(key, conversation);
    conversation.push({ role: "assistant", content: [{ type: "text", text: answer }] });
    els.answer.textContent = answer;
    els.followupQuestion.textContent = "";
    els.followupQuestion.classList.add("hidden");
    answerSlideId = sourceSlide.id;
    document.dispatchEvent(new CustomEvent("answer-ready", { detail: { slideId: answerSlideId } }));
    els.followupButton.disabled = false;
    const answerLabel = correctAnswer ? `；本页标准答案 ${correctAnswer}` : "；本页未标注标准答案";
    const completed = followup ? "课堂追问已回答" : "讲解完成";
    setStatus(rainStats ? `${completed}${answerLabel}；已使用雨课堂作答分布。` : `${completed}${answerLabel}；未使用作答分布：${rainResult.reason}。`, false);
  } catch (error) {
    document.dispatchEvent(new Event("answer-failed"));
    els.answer.textContent = error.message || String(error);
    setStatus(error.message || String(error), true);
  } finally {
    isBusy = false;
    els.explain.disabled = false;
    els.followupButton.disabled = conversation.length === 0;
    els.followupButton.textContent = "发送";
  }
}

async function askFollowup() {
  if (isBusy) return;
  const text = els.followupInput.value.trim();
  const key = localStorage.getItem(STORAGE_KEY);
  if (!text) return setStatus("请先输入追问问题。", true);
  if (!key) return setStatus("请先在设置中保存 AppKey。", true);
  if (conversation.length === 0) return setStatus("请先讲解本题，再继续追问。", true);
  const messages = [...conversation, { role: "user", content: [{ type: "text", text }] }];
  els.followupQuestion.textContent = "你的追问：" + text;
  els.followupQuestion.classList.remove("hidden");
  isBusy = true;
  document.dispatchEvent(new Event("answer-pending"));
  setBusy(true, "正在回答追问…");
  try {
    const answer = await callYuanqi(key, messages);
    conversation = [...messages, { role: "assistant", content: [{ type: "text", text: answer }] }];
    if (els.followupInput.value.trim() === text) els.followupInput.value = "";
    els.answer.textContent = answer;
    document.dispatchEvent(new CustomEvent("answer-ready", { detail: { slideId: answerSlideId } }));
    setStatus("追问完成。", false);
  } catch (error) {
    document.dispatchEvent(new Event("answer-failed"));
    setStatus(error.message || String(error), true);
  } finally {
    isBusy = false;
    els.explain.disabled = false;
    els.followupButton.disabled = conversation.length === 0;
    els.followupButton.textContent = "发送";
  }
}

function setBusy(disabled, message) {
  els.explain.disabled = disabled;
  els.followupButton.disabled = disabled;
  els.followupButton.textContent = disabled ? "请稍候…" : "发送";
  setStatus(message, false);
}

async function callYuanqi(key, messages) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 90000);
  try {
  const response = await fetch(API_URL, {
    signal: controller.signal,
    method: "POST",
    headers: {
      "Authorization": `Bearer ${key}`,
      "Content-Type": "application/json",
      "X-Source": "openapi",
    },
    body: JSON.stringify({
      assistant_id: ASSISTANT_ID,
      user_id: "ppt-classroom",
      stream: false,
      messages,
    }),
  });
  const payload = await response.json().catch(() => null);
  if (!response.ok) throw new Error(payload?.message || `腾讯元器调用失败（${response.status}）`);
  const answer = payload?.choices?.[0]?.message?.content;
  if (!answer) throw new Error("腾讯元器没有返回可显示的答案。");
  return typeof answer === "string" ? answer : answer.map((part) => part.text || "").join("");
  } catch (error) {
    if (error.name === "AbortError") throw new Error("回答等待超时，请再次发送；追问内容已保留。");
    throw error;
  } finally { clearTimeout(timeout); }
}
