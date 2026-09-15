const API_URL = "https://yuanqi.tencent.com/openapi/v1/agent/chat/completions";
const ASSISTANT_ID = "2096407988983757888";
const STORAGE_KEY = "yuanqi-app-key";

const els = {};
let conversation = [];

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
    followupInput: document.getElementById("followupInput"),
    followupButton: document.getElementById("followupButton"),
  });

  els.settingsButton.addEventListener("click", () => els.settings.classList.toggle("hidden"));
  els.saveKey.addEventListener("click", saveKey);
  els.clearKey.addEventListener("click", clearKey);
  els.explain.addEventListener("click", explainCurrentSlide);
  els.followupButton.addEventListener("click", askFollowup);
  els.followupInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") askFollowup();
  });

  if (info.host !== Office.HostType.PowerPoint) {
    setStatus("请在 PowerPoint 中打开此加载项。", true);
    return;
  }

  els.explain.disabled = false;
  setStatus("已就绪：切到题目页后点击一次按钮。", false);
  if (!localStorage.getItem(STORAGE_KEY)) els.settings.classList.remove("hidden");
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

function getCurrentSlideIndex() {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.SlideRange,
      (result) => {
        if (
          result.status === Office.AsyncResultStatus.Succeeded &&
          result.value?.slides?.length
        ) {
          resolve(result.value.slides[0].index - 1);
        } else {
          resolve(null);
        }
      }
    );
  });
}

async function readCurrentSlideText() {
  const currentSlideIndex = await getCurrentSlideIndex();
  return PowerPoint.run(async (context) => {
    let slide;
    if (currentSlideIndex !== null) {
      slide = context.presentation.slides.getItemAt(currentSlideIndex);
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
  const key = localStorage.getItem(STORAGE_KEY);
  if (!key) {
    els.settings.classList.remove("hidden");
    return setStatus("首次使用：请粘贴并保存 AppKey。", true);
  }

  setBusy(true, "正在读取当前页…");
  try {
    const question = await readCurrentSlideText();
    if (!question) throw new Error("当前页没有读取到文字。图片题请把题目文字放在一个文本框中。");

    const prompt = `你是大学宏观经济学课堂的AI课程助教。请讲解下面的题目：\n\n${question}\n\n要求：先明确给出正确答案；再解释核心原理；有选项时逐项判断；不虚构题目中没有的数据；控制在课堂60—90秒可讲完；最后用一句话总结考点。`;
    conversation = [{ role: "user", content: [{ type: "text", text: prompt }] }];
    const answer = await callYuanqi(key, conversation);
    conversation.push({ role: "assistant", content: [{ type: "text", text: answer }] });
    els.answer.textContent = answer;
    els.followupButton.disabled = false;
    setStatus("讲解完成。", false);
  } catch (error) {
    setStatus(error.message || String(error), true);
  } finally {
    els.explain.disabled = false;
  }
}

async function askFollowup() {
  const text = els.followupInput.value.trim();
  const key = localStorage.getItem(STORAGE_KEY);
  if (!text || !key || conversation.length === 0) return;
  els.followupInput.value = "";
  setBusy(true, "正在回答追问…");
  try {
    conversation.push({ role: "user", content: [{ type: "text", text }] });
    const answer = await callYuanqi(key, conversation);
    conversation.push({ role: "assistant", content: [{ type: "text", text: answer }] });
    els.answer.textContent = answer;
    setStatus("追问完成。", false);
  } catch (error) {
    setStatus(error.message || String(error), true);
  } finally {
    els.explain.disabled = false;
  }
}

function setBusy(disabled, message) {
  els.explain.disabled = disabled;
  els.followupButton.disabled = disabled;
  setStatus(message, false);
}

async function callYuanqi(key, messages) {
  const response = await fetch(API_URL, {
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
}
