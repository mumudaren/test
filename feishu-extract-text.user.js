// ==UserScript==
// @name         飞书文档提取纯文字
// @namespace    feishu-extract-text
// @version      1.0.0
// @description  在已打开的飞书新版文档或知识库文档上，提取全文纯文字，便于复制和打印
// @author       you
// @match        https://*.feishu.cn/wiki/*
// @match        https://*.feishu.cn/docx/*
// @match        https://*.larksuite.com/wiki/*
// @match        https://*.larksuite.com/docx/*
// @match        https://*.larkoffice.com/wiki/*
// @match        https://*.larkoffice.com/docx/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const DOCX_TYPE = 22;

  function tokenFromUrl() {
    const match = location.pathname.match(/\/(wiki|docx)\/([A-Za-z0-9]+)/);
    if (!match) return null;
    return { kind: match[1], token: match[2] };
  }

  async function readJson(url) {
    const response = await fetch(url, { credentials: "include" });
    if (!response.ok) {
      throw new Error("请求失败 " + response.status);
    }
    const json = await response.json();
    if (json.code !== 0) {
      throw new Error(json.msg || "飞书接口返回失败");
    }
    return json.data;
  }

  async function resolveDoc(target) {
    if (target.kind === "docx") {
      return { id: target.token, title: document.title.replace(/\s*-\s*飞书云文档\s*$/, "") };
    }
    const data = await readJson(
      location.origin +
        "/space/api/wiki/v2/tree/get_node/?wiki_token=" +
        encodeURIComponent(target.token)
    );
    if (data.obj_type !== DOCX_TYPE) {
      throw new Error("这一页不是新版文档。目前只支持知识库里的文档和 /docx/ 链接。");
    }
    return { id: data.obj_token, title: data.title || document.title };
  }

  async function fetchBlocks(docId) {
    const blockMap = {};
    let cursor = "";
    for (let page = 0; page < 40; page += 1) {
      const url = new URL("/space/api/docx/pages/client_vars", location.origin);
      url.searchParams.set("id", docId);
      url.searchParams.set("limit", "500");
      if (cursor) url.searchParams.set("cursor", cursor);
      const data = await readJson(url.toString());
      Object.assign(blockMap, data.block_map || {});
      if (!data.has_more || !data.cursor || data.cursor === cursor) break;
      cursor = data.cursor;
    }
    return blockMap;
  }

  function blockText(data) {
    const pieces = data && data.text && data.text.initialAttributedTexts && data.text.initialAttributedTexts.text;
    if (!pieces) return "";
    return Object.keys(pieces)
      .sort((a, b) => Number(a) - Number(b))
      .map((key) => pieces[key])
      .join("")
      .replace(/\u200b/g, "")
      .trim();
  }

  function render(blockMap, docId) {
    const page = blockMap[docId] || Object.values(blockMap).find((block) => block.data && block.data.type === "page");
    if (!page) throw new Error("没有找到文档正文。请等页面完全打开后再试。");

    function walk(id) {
      const block = blockMap[id];
      if (!block || !block.data) return [];
      const data = block.data;
      const type = data.type;
      const text = blockText(data);

      if (type === "page" || type === "callout" || type === "quote_container" || type === "grid" || type === "grid_column" || type === "table_cell") {
        return (data.children || []).flatMap(walk);
      }

      if (type === "table") {
        const lines = [];
        for (const rowId of data.rows_id || []) {
          for (const columnId of data.columns_id || []) {
            const cell = data.cell_set && data.cell_set[rowId + columnId];
            if (cell) lines.push(...walk(cell.block_id));
          }
        }
        return lines;
      }

      if (type === "image") return ["【图片】"];

      if (type === "bullet" || type === "ordered") {
        const nested = (data.children || []).flatMap(walk);
        return [...(text ? ["• " + text] : []), ...nested];
      }

      if (/^heading[1-6]$/.test(type)) {
        return text ? ["", text, ""] : [];
      }

      const nested = (data.children || []).flatMap(walk);
      return [...(text ? [text] : []), ...nested];
    }

    const collapsed = [];
    for (const line of walk(page.id || docId)) {
      if (line === "【图片】" && collapsed[collapsed.length - 1] && collapsed[collapsed.length - 1].startsWith("【图片")) {
        const previous = collapsed[collapsed.length - 1];
        const count = previous === "【图片】" ? 2 : Number(previous.match(/×(\d+)/)[1]) + 1;
        collapsed[collapsed.length - 1] = "【图片 ×" + count + "】";
      } else {
        collapsed.push(line);
      }
    }

    const output = [];
    for (const line of collapsed) {
      if (line === "" && (output.length === 0 || output[output.length - 1] === "")) continue;
      output.push(line);
    }
    return output.join("\n").trim() + "\n";
  }

  function safeFilename(title) {
    const name = (title || "飞书文档").replace(/[\\/:*?"<>|]/g, " ").replace(/\s+/g, " ").trim();
    return (name || "飞书文档") + ".txt";
  }

  function buildPanel() {
    const style = document.createElement("style");
    style.textContent = [
      "#feishu-extract-btn{position:fixed;right:20px;bottom:24px;z-index:2147483646;border:0;border-radius:999px;padding:10px 16px;background:#1456f0;color:#fff;font-size:14px;cursor:pointer;box-shadow:0 6px 18px rgba(20,86,240,.28)}",
      "#feishu-extract-panel{position:fixed;right:20px;bottom:72px;z-index:2147483646;width:min(560px,calc(100vw - 32px));height:min(70vh,640px);display:flex;flex-direction:column;background:#fff;color:#1f2329;border:1px solid #dee0e3;border-radius:12px;box-shadow:0 12px 40px rgba(31,35,41,.18);font:14px/1.5 -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}",
      "#feishu-extract-panel header{display:flex;align-items:center;gap:8px;padding:12px 14px;border-bottom:1px solid #eff0f1}",
      "#feishu-extract-panel header strong{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}",
      "#feishu-extract-panel header button,#feishu-extract-actions button{border:1px solid #dee0e3;background:#fff;border-radius:8px;padding:6px 10px;cursor:pointer}",
      "#feishu-extract-status{padding:8px 14px;color:#646a73;font-size:12px}",
      "#feishu-extract-text{flex:1;margin:0 14px;border:1px solid #dee0e3;border-radius:8px;padding:10px;resize:none;font:13px/1.6 ui-monospace,SFMono-Regular,Menlo,Consolas,monospace}",
      "#feishu-extract-actions{display:flex;gap:8px;padding:12px 14px}",
    ].join("");
    document.documentElement.appendChild(style);

    const button = document.createElement("button");
    button.id = "feishu-extract-btn";
    button.type = "button";
    button.textContent = "提取文字";
    document.documentElement.appendChild(button);

    const panel = document.createElement("section");
    panel.id = "feishu-extract-panel";
    panel.hidden = true;
    panel.innerHTML = [
      "<header><strong>提取文字</strong><button type='button' data-act='close'>关闭</button></header>",
      "<div id='feishu-extract-status'>还没有提取</div>",
      "<textarea id='feishu-extract-text' readonly></textarea>",
      "<div id='feishu-extract-actions'>",
      "<button type='button' data-act='copy'>复制全文</button>",
      "<button type='button' data-act='download'>下载 txt</button>",
      "</div>",
    ].join("");
    document.documentElement.appendChild(panel);

    const status = panel.querySelector("#feishu-extract-status");
    const textarea = panel.querySelector("#feishu-extract-text");
    let currentTitle = "飞书文档";

    async function extract() {
      panel.hidden = false;
      status.textContent = "正在读取正文…";
      textarea.value = "";
      const target = tokenFromUrl();
      if (!target) {
        status.textContent = "当前网址不是飞书文档或知识库页面。";
        return;
      }
      try {
        const doc = await resolveDoc(target);
        currentTitle = doc.title;
        panel.querySelector("header strong").textContent = doc.title;
        const blocks = await fetchBlocks(doc.id);
        const text = render(blocks, doc.id);
        textarea.value = text;
        status.textContent = "已提取 " + text.length + " 个字符。图片位置用【图片】标出。";
      } catch (error) {
        status.textContent = error && error.message ? error.message : "提取失败";
      }
    }

    button.addEventListener("click", extract);
    panel.addEventListener("click", async (event) => {
      const action = event.target && event.target.getAttribute("data-act");
      if (action === "close") panel.hidden = true;
      if (action === "copy") {
        if (!textarea.value) return;
        await navigator.clipboard.writeText(textarea.value);
        status.textContent = "已复制到剪贴板。";
      }
      if (action === "download") {
        if (!textarea.value) return;
        const blob = new Blob([textarea.value], { type: "text/plain;charset=utf-8" });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = safeFilename(currentTitle);
        link.click();
        URL.revokeObjectURL(link.href);
      }
    });
  }

  buildPanel();
})();
