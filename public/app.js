const quill = new Quill("#editor", {
  theme: "snow",
  modules: {
    toolbar: [["bold", "italic", "underline"], [{ background: [] }]]
  }
});

let currentDocxB64 = null;
let currentTxt = null;

// ---------- tab switching ----------
document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    document
      .querySelectorAll(".tab")
      .forEach((t) => t.classList.remove("active"));
    document
      .querySelectorAll(".tab-content")
      .forEach((c) => c.classList.add("hidden"));
    tab.classList.add("active");
    document
      .getElementById("tab-" + tab.dataset.tab)
      .classList.remove("hidden");
  });
});

// ---------- drag & drop for upload ----------
const fileInput = document.getElementById("file-input");
const uploadZone = document.getElementById("tab-upload");
uploadZone.addEventListener("dragover", (e) => {
  e.preventDefault();
  uploadZone.classList.add("dragover");
});
uploadZone.addEventListener("dragleave", () =>
  uploadZone.classList.remove("dragover")
);
uploadZone.addEventListener("drop", (e) => {
  e.preventDefault();
  uploadZone.classList.remove("dragover");
  if (e.dataTransfer.files.length) {
    fileInput.files = e.dataTransfer.files;
  }
});

// ---------- process ----------
document.getElementById("process-btn").addEventListener("click", async () => {
  const activeTab = document.querySelector(".tab.active").dataset.tab;
  const fd = new FormData();

  if (activeTab === "upload") {
    const file = fileInput.files[0];
    if (!file) return setStatus("Please select a .docx file", "error");
    if (!file.name.toLowerCase().endsWith(".docx"))
      return setStatus("Only .docx files are supported", "error");
    fd.append("file", file);
  } else {
    const html = quill.root.innerHTML;
    if (!html || html.trim() === "<p><br></p>")
      return setStatus("Please paste some content", "error");
    fd.append("html", html);
  }

  fd.append("sep", document.getElementById("separator").value);
  const settings = readSettings();
  for (const [k, v] of Object.entries(settings)) fd.append(k, v);

  setStatus("Processing…", "info");
  const btn = document.getElementById("process-btn");
  btn.disabled = true;

  try {
    const res = await fetch("/api/flatten", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Processing failed");

    currentDocxB64 = data.docx_b64;
    currentTxt = data.txt;

    renderOutput(data.blocks, data.sep);
    setStatus("✓ Done — download buttons at the top", "success");
  } catch (e) {
    setStatus("Error: " + e.message, "error");
  } finally {
    btn.disabled = false;
  }
});

function readSettings() {
  const v = (id) => document.getElementById(id).value;
  const c = (id) => document.getElementById(id).checked;
  return {
    font_name: v("font_name"),
    font_size: v("font_size"),
    table_style: v("table_style"),
    autofit_mode: v("autofit_mode"),
    space_before: v("space_before"),
    space_after: v("space_after"),
    line_spacing: v("line_spacing"),
    margin_top: v("margin_top"),
    margin_bottom: v("margin_bottom"),
    margin_left: v("margin_left"),
    margin_right: v("margin_right"),
    gutter: v("gutter"),
    gutter_position: v("gutter_position"),
    num_cols: v("num_cols"),
    col_spacing: v("col_spacing"),
    col_line: c("col_line") ? "true" : "false"
  };
}

// ---------- rendering ----------
const COLOR_MAP = {
  YELLOW: "#FFFF00",
  BRIGHT_GREEN: "#00FF00",
  GREEN: "#00FF00",
  CYAN: "#00FFFF",
  TURQUOISE: "#00FFFF",
  MAGENTA: "#FF00FF",
  PINK: "#FF00FF",
  VIOLET: "#FF00FF",
  BLUE: "#0000FF",
  RED: "#FF0000",
  DARK_BLUE: "#000080",
  TEAL: "#008080",
  DARK_GREEN: "#008000",
  DARK_RED: "#800000",
  DARK_YELLOW: "#808000",
  GRAY_25: "#C0C0C0",
  GRAY_50: "#808080",
  BLACK: "#000000"
};

function colorToCss(value) {
  if (value.startsWith("#")) return value;
  return COLOR_MAP[value.toUpperCase()] || "#FFFF00";
}

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function renderTextWithHighlights(text) {
  const parts = text.split(/(<<HL:[^>]*>>|<<\/HL>>)/);
  return parts
    .map((part) => {
      const openMatch = part.match(/^<<HL:([^>]*)>>$/);
      if (openMatch) {
        return `<span class="hl" style="background-color:${colorToCss(openMatch[1])}">`;
      }
      if (part === "<</HL>>") return "</span>";
      return escapeHtml(part);
    })
    .join("");
}

function stripHl(s) {
  return s.replace(/<<HL:[^>]*>>|<<\/HL>>/g, "");
}

function renderOutput(blocks, sep) {
  const output = document.getElementById("output-section");
  output.classList.remove("hidden");

  const preview = document.getElementById("preview");
  preview.innerHTML = "";

  blocks.forEach((block) => {
    if (block.kind === "text") {
      const text = block.data.join(sep);
      const div = document.createElement("div");
      div.innerHTML = renderTextWithHighlights(text);
      div.style.marginBottom = "12px";
      preview.appendChild(div);
    } else {
      const table = document.createElement("table");
      table.className = "preview-table";
      block.data.forEach((row, i) => {
        const tr = document.createElement("tr");
        row.forEach((cell) => {
          const td = document.createElement(i === 0 ? "th" : "td");
          td.innerHTML = renderTextWithHighlights(cell);
          tr.appendChild(td);
        });
        table.appendChild(tr);
      });
      preview.appendChild(table);
    }
  });

  const raw = blocks
    .filter((b) => b.kind === "text")
    .map((b) => stripHl(b.data.join(sep)))
    .join("\n");
  document.getElementById("raw-text").textContent = raw || "(no text blocks)";

  // Scroll to output — download buttons are at the top of the visible area
  output.scrollIntoView({ behavior: "smooth", block: "start" });
}

// ---------- downloads ----------
document.getElementById("download-docx").addEventListener("click", () => {
  if (!currentDocxB64) return;
  const bin = atob(currentDocxB64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
  triggerDownload(blob, "flattened.docx");
});

document.getElementById("download-txt").addEventListener("click", () => {
  if (!currentTxt) return;
  triggerDownload(
    new Blob([currentTxt], { type: "text/plain" }),
    "flattened.txt"
  );
});

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function setStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status " + (type || "");
}
