const API_URL = "https://yuanqi.tencent.com/openapi/v1/agent/chat/completions";
const ASSISTANT_ID = "2096407988983757888";
const STORAGE_KEY = "yuanqi-app-key";
const AUTO_SLIDES_KEY = "ppt-ai-auto-slide-ids";

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
    followupInput: document.getElementById("followupInput"),
    followupButton: document.getElementById("followupButton"),
    followupQuestion: document.getElementById("followupQuestion"),
  });

  els.settingsButton.addEventListener("click", () => els.settings.classList.toggle("hidden"));
  els.saveKey.addEventListener("click", saveKey);
  els.clearKey.addEventListener("click", clearKey);
  els.toggleAuto.addEventListener("click", toggleCurrentSlideAuto);
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
  setBusy(true, "正在读取当前页…");
  try {
    const sourceSlide = await getCurrentSlideInfo();
    if (!sourceSlide) throw new Error("没有识别到正在显示的幻灯片，请重试。");
    const question = await readCurrentSlideText(sourceSlide);
    if (!question) throw new Error("当前页没有读取到文字。图片题请把题目文字放在一个文本框中。");

    const prompt = `你是大学宏观经济学课堂的AI课程助教。请讲解下面的题目：\n\n${question}\n\n要求：先明确给出正确答案；再解释核心原理；有选项时逐项判断；不虚构题目中没有的数据；控制在课堂60—90秒可讲完；最后用一句话总结考点。`;
    conversation = [{ role: "user", content: [{ type: "text", text: prompt }] }];
    const answer = await callYuanqi(key, conversation);
    conversation.push({ role: "assistant", content: [{ type: "text", text: answer }] });
    els.answer.textContent = answer;
    els.followupQuestion.textContent = "";
    els.followupQuestion.classList.add("hidden");
    answerSlideId = sourceSlide.id;
    document.dispatchEvent(new CustomEvent("answer-ready", { detail: { slideId: answerSlideId } }));
    els.followupButton.disabled = false;
    setStatus("讲解完成。", false);
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
