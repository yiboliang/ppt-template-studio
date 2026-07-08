const state = {
  themes: [],
  layouts: [],
  layoutMap: {},
  slides: [],
};

const el = (id) => document.getElementById(id);

async function loadMeta() {
  const [themes, layouts] = await Promise.all([
    fetch("/api/themes").then((r) => r.json()),
    fetch("/api/layouts").then((r) => r.json()),
  ]);
  state.themes = themes;
  state.layouts = layouts;
  layouts.forEach((l) => (state.layoutMap[l.id] = l));

  const themeSelect = el("themeSelect");
  themeSelect.innerHTML = themes
    .map((t) => `<option value="${t.id}">${t.display_name}</option>`)
    .join("");

  const layoutSelect = el("layoutSelect");
  layoutSelect.innerHTML = layouts
    .map((l) => `<option value="${l.id}">${l.name}</option>`)
    .join("");

  renderSlideForm();
  layoutSelect.addEventListener("change", renderSlideForm);
}

function renderSlideForm() {
  const layoutId = el("layoutSelect").value;
  const layout = state.layoutMap[layoutId];
  el("layoutHint").textContent = layout.description || "";

  const form = el("slideForm");
  form.innerHTML = layout.placeholders
    .map((p) => {
      const req = p.required ? " *必填" : "";
      const labelText = `${p.key}${req}`;
      if (p.type === "bullet_list" || p.type === "numbered_list") {
        return `
          <label>${labelText}（每行一条）</label>
          <textarea data-key="${p.key}" data-type="${p.type}" rows="5"></textarea>`;
      }
      return `
        <label>${labelText}</label>
        <input type="text" data-key="${p.key}" data-type="${p.type}" />`;
    })
    .join("");
}

function collectFormValues() {
  const layoutId = el("layoutSelect").value;
  const content = {};
  document.querySelectorAll("#slideForm [data-key]").forEach((field) => {
    const key = field.dataset.key;
    const type = field.dataset.type;
    const raw = field.value.trim();
    if (!raw) return;
    if (type === "bullet_list" || type === "numbered_list") {
      content[key] = raw.split("\n").map((s) => s.trim()).filter(Boolean);
    } else {
      content[key] = raw;
    }
  });
  return { layout: layoutId, content };
}

function renderOutline() {
  const list = el("slideList");
  el("slideCount").textContent = `(${state.slides.length})`;
  list.innerHTML = state.slides
    .map((s, i) => {
      const layout = state.layoutMap[s.layout];
      const titleLike = s.content.title || s.content.quote || s.content.index || "(无标题)";
      return `
      <li class="slide-item">
        <div class="meta">
          <span>#${i + 1} · ${layout ? layout.name : s.layout}</span>
          <span class="row-btns">
            <button data-act="up" data-idx="${i}">↑</button>
            <button data-act="down" data-idx="${i}">↓</button>
            <button data-act="del" data-idx="${i}">删除</button>
          </span>
        </div>
        <div class="title">${escapeHtml(String(titleLike))}</div>
      </li>`;
    })
    .join("");
  updateJsonBox();
}

function escapeHtml(str) {
  return str.replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function updateJsonBox() {
  const deck = { theme: el("themeSelect").value, slides: state.slides };
  el("jsonBox").value = JSON.stringify(deck, null, 2);
}

function setStatus(msg, isError = false) {
  const box = el("statusMsg");
  box.textContent = msg;
  box.style.color = isError ? "#ff6b6b" : "";
}

el("addSlideBtn").addEventListener("click", () => {
  const slide = collectFormValues();
  state.slides.push(slide);
  renderOutline();
  setStatus(`已添加第 ${state.slides.length} 页`);
});

el("slideList").addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-act]");
  if (!btn) return;
  const idx = Number(btn.dataset.idx);
  const act = btn.dataset.act;
  if (act === "del") state.slides.splice(idx, 1);
  if (act === "up" && idx > 0) [state.slides[idx - 1], state.slides[idx]] = [state.slides[idx], state.slides[idx - 1]];
  if (act === "down" && idx < state.slides.length - 1) [state.slides[idx + 1], state.slides[idx]] = [state.slides[idx], state.slides[idx + 1]];
  renderOutline();
});

el("clearBtn").addEventListener("click", () => {
  state.slides = [];
  renderOutline();
  setStatus("已清空");
});

el("themeSelect").addEventListener("change", updateJsonBox);

el("applyJsonBtn").addEventListener("click", () => {
  try {
    const parsed = JSON.parse(el("jsonBox").value);
    if (!parsed.theme || !Array.isArray(parsed.slides)) {
      throw new Error("JSON 需要包含 'theme' 字符串和 'slides' 数组");
    }
    el("themeSelect").value = parsed.theme;
    state.slides = parsed.slides;
    renderOutline();
    setStatus("已从 JSON 应用到大纲");
  } catch (err) {
    setStatus(`JSON 解析失败: ${err.message}`, true);
  }
});

el("generateBtn").addEventListener("click", async () => {
  if (state.slides.length === 0) {
    setStatus("大纲为空，请先添加幻灯片", true);
    return;
  }
  setStatus("正在生成 PPT ...");
  const deck = { theme: el("themeSelect").value, slides: state.slides };
  try {
    const resp = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(deck),
    });
    if (!resp.ok) {
      const err = await resp.json().catch(() => ({ detail: resp.statusText }));
      throw new Error(err.detail || "生成失败");
    }
    const blob = await resp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const cd = resp.headers.get("Content-Disposition") || "";
    const match = cd.match(/filename="?([^"]+)"?/);
    a.href = url;
    a.download = match ? match[1] : "deck.pptx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    setStatus("生成成功，已开始下载 ✅");
  } catch (err) {
    setStatus(`生成失败: ${err.message}`, true);
  }
});

loadMeta();
