// ==UserScript==
// @name         飞书文档提取并打印
// @namespace    feishu-extract-text
// @version      1.1.0
// @description  把已打开的飞书新版文档导出成带图片、可打印的网页，尽量保持标题、分栏和表格位置
// @match        https://*.feishu.cn/wiki/*
// @match        https://*.feishu.cn/docx/*
// @match        https://*.larksuite.com/wiki/*
// @match        https://*.larksuite.com/docx/*
// @match        https://*.larkoffice.com/wiki/*
// @match        https://*.larkoffice.com/docx/*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      feishu.cn
// @connect      larksuite.com
// @connect      larkoffice.com
// @connect      feishucdn.com
// @connect      internal-api-drive-stream.feishu.cn
// ==/UserScript==

(function () {
  "use strict";

  const DOCX_TYPE = 22;
  const pageFetch = (unsafeWindow || window).fetch.bind(unsafeWindow || window);

  function tokenFromUrl() {
    const match = location.pathname.match(/\/(wiki|docx)\/([A-Za-z0-9]+)/);
    if (!match) return null;
    return { kind: match[1], token: match[2] };
  }

  async function readJson(url) {
    const response = await pageFetch(url, { credentials: "include" });
    if (!response.ok) throw new Error("请求失败 " + response.status);
    const json = await response.json();
    if (json.code !== 0 && json.code !== undefined) throw new Error(json.msg || json.message || "飞书接口返回失败");
    return json.data;
  }

  async function resolveDoc(target) {
    if (target.kind === "docx") {
      return { id: target.token, title: document.title.replace(/\s*-\s*飞书云文档\s*$/, "") };
    }
    const data = await readJson(
      location.origin + "/space/api/wiki/v2/tree/get_node/?wiki_token=" + encodeURIComponent(target.token)
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

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function gmGetBlob(url) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: "GET",
        url,
        responseType: "blob",
        onload: (res) => {
          const blob = res.response;
          const type = blob && blob.type ? blob.type : "";
          if (res.status >= 200 && res.status < 300 && blob && blob.size > 200 && !type.includes("json") && !type.includes("text")) {
            resolve(blob);
            return;
          }
          reject(new Error("图片下载失败 " + res.status));
        },
        onerror: () => reject(new Error("图片网络错误")),
      });
    });
  }

  function blobToDataUrl(blob) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(new Error("图片读取失败"));
      reader.readAsDataURL(blob);
    });
  }

  function imageUrls(token, blockId) {
    const mount = "mount_point=docx_image&mount_node_token=" + encodeURIComponent(blockId);
    return [
      location.origin + "/space/api/box/stream/download/all/" + token + "/?" + mount,
      "https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/all/" + token + "/?" + mount,
      "https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/preview/" + token + "?preview_type=16",
      "https://internal-api-drive-stream.feishu.cn/space/api/box/stream/download/v2/cover/" +
        token +
        "/?fallback_source=1&height=1920&policy=equal&" +
        mount,
    ];
  }

  async function downloadImage(token, blockId) {
    let lastError = null;
    for (const url of imageUrls(token, blockId)) {
      try {
        return await blobToDataUrl(await gmGetBlob(url));
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError || new Error("图片下载失败");
  }

  function collectImages(blockMap, docId) {
    const images = [];
    const seen = new Set();
    (function walk(id) {
      const block = blockMap[id];
      if (!block || !block.data || seen.has(id)) return;
      seen.add(id);
      const data = block.data;
      if (data.type === "image" && data.image && data.image.token) {
        images.push({ id, token: data.image.token, name: data.image.name || "image" });
      }
      const childIds = data.children ? data.children.slice() : [];
      if (data.type === "table" && data.cell_set) {
        Object.values(data.cell_set).forEach((cell) => {
          if (cell && cell.block_id) childIds.push(cell.block_id);
        });
      }
      childIds.forEach(walk);
    })(docId);
    return images;
  }

  async function loadImages(images, onProgress) {
    const result = {};
    let done = 0;
    const queue = images.slice();
    async function worker() {
      while (queue.length) {
        const image = queue.shift();
        try {
          result[image.id] = await downloadImage(image.token, image.id);
        } catch (error) {
          result[image.id] = "";
        }
        done += 1;
        onProgress(done, images.length);
      }
    }
    await Promise.all([worker(), worker(), worker()]);
    return result;
  }

  function renderHtml(blockMap, docId, title, images) {
    const page = blockMap[docId] || Object.values(blockMap).find((block) => block.data && block.data.type === "page");
    if (!page) throw new Error("没有找到文档正文。请等页面完全打开后再试。");

    function childrenHtml(ids) {
      return (ids || []).map(renderBlock).join("");
    }

    function renderBlock(id) {
      const block = blockMap[id];
      if (!block || !block.data) return "";
      const data = block.data;
      const type = data.type;
      const text = blockText(data);

      if (type === "page") return childrenHtml(data.children);
      if (type === "callout") {
        const background = data.background_color ? "background:" + data.background_color + ";" : "";
        const border = data.border_color ? "border-color:" + data.border_color + ";" : "";
        return '<div class="callout" style="' + background + border + '">' + childrenHtml(data.children) + "</div>";
      }
      if (type === "quote_container") return "<blockquote>" + childrenHtml(data.children) + "</blockquote>";
      if (type === "grid") return '<div class="grid">' + childrenHtml(data.children) + "</div>";
      if (type === "grid_column") {
        const width = Math.round((data.width_ratio || 1) * 1000) / 10;
        return '<div class="grid-col" style="flex-basis:' + width + '%">' + childrenHtml(data.children) + "</div>";
      }
      if (type === "table_cell") return childrenHtml(data.children);
      if (type === "table") {
        const columns = data.columns_id || [];
        const total = columns.reduce((sum, columnId) => sum + ((data.column_set && data.column_set[columnId] && data.column_set[columnId].column_width) || 1), 0);
        const colgroup = columns
          .map((columnId) => {
            const width = (data.column_set && data.column_set[columnId] && data.column_set[columnId].column_width) || 1;
            return '<col style="width:' + (width / total) * 100 + '%">';
          })
          .join("");
        const rows = (data.rows_id || [])
          .map((rowId) => {
            const cells = columns
              .map((columnId) => {
                const cell = data.cell_set && data.cell_set[rowId + columnId];
                return "<td>" + (cell ? renderBlock(cell.block_id) : "") + "</td>";
              })
              .join("");
            return "<tr>" + cells + "</tr>";
          })
          .join("");
        return "<table><colgroup>" + colgroup + "</colgroup><tbody>" + rows + "</tbody></table>";
      }
      if (type === "image") {
        const src = images[id];
        const name = escapeHtml((data.image && data.image.name) || "图片");
        if (!src) return '<p class="missing">图片未能下载：' + name + "</p>";
        return '<figure><img src="' + src + '" alt="' + name + '"></figure>';
      }
      if (type === "bullet" || type === "ordered") {
        return '<p class="bullet">' + escapeHtml(text) + "</p>" + childrenHtml(data.children);
      }
      if (/^heading[1-6]$/.test(type)) {
        const level = type.slice(7);
        return text ? "<h" + level + ">" + escapeHtml(text) + "</h" + level + ">" : "";
      }
      const body = text ? "<p>" + escapeHtml(text).replace(/\n/g, "<br>") + "</p>" : "";
      return body + childrenHtml(data.children);
    }

    const body = renderBlock(page.id || docId);
    return [
      "<!DOCTYPE html><html lang='zh-CN'><head><meta charset='utf-8'>",
      "<title>",
      escapeHtml(title),
      "</title><style>",
      "@page{size:A4;margin:14mm}",
      "body{margin:0 auto;max-width:820px;padding:24px 28px 48px;color:#1f2329;font:15px/1.7 'PingFang SC','Microsoft YaHei',sans-serif}",
      "h1,h2,h3,h4,h5,h6{line-height:1.4;page-break-after:avoid}",
      "h1{font-size:26px}h2{font-size:22px}h3{font-size:18px}h4{font-size:16px}",
      "p{margin:6px 0}p.bullet{margin:2px 0 2px 1.2em;text-indent:-1.2em}p.bullet:before{content:'• ';}",
      ".callout{margin:10px 0;padding:10px 14px;border:1px solid #f3c07a;border-radius:8px;background:#fff8ef}",
      "blockquote{margin:10px 0;padding-left:12px;border-left:3px solid #bbb;color:#444}",
      ".grid{display:flex;gap:16px;align-items:flex-start}",
      ".grid-col{min-width:0}",
      "table{width:100%;border-collapse:collapse;margin:10px 0;page-break-inside:auto}",
      "td{vertical-align:top;padding:6px 8px;border:1px solid #ececec}",
      "figure{margin:8px 0}img{max-width:100%;height:auto;display:block}",
      ".missing{color:#a15c07;font-size:13px}",
      ".toolbar{position:sticky;top:0;background:#fff;padding:8px 0 12px;border-bottom:1px solid #eee;margin-bottom:16px}",
      ".toolbar button{border:0;background:#1456f0;color:#fff;border-radius:8px;padding:8px 14px;cursor:pointer}",
      "@media print{.toolbar{display:none}body{max-width:none;padding:0}}",
      "</style></head><body>",
      "<div class='toolbar'><button onclick='print()'>打印</button></div>",
      body,
      "</body></html>",
    ].join("");
  }

  function safeFilename(title) {
    const name = (title || "飞书文档").replace(/[\\/:*?"<>|]/g, " ").replace(/\s+/g, " ").trim();
    return (name || "飞书文档") + ".html";
  }

  function downloadHtml(filename, html) {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(new Blob([html], { type: "text/html;charset=utf-8" }));
    link.download = filename;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  function buildPanel() {
    const style = document.createElement("style");
    style.textContent = [
      "#feishu-extract-btn{position:fixed;right:20px;bottom:24px;z-index:2147483646;border:0;border-radius:999px;padding:10px 16px;background:#1456f0;color:#fff;font-size:14px;cursor:pointer;box-shadow:0 6px 18px rgba(20,86,240,.28)}",
      "#feishu-extract-panel{position:fixed;right:20px;bottom:72px;z-index:2147483646;width:min(420px,calc(100vw - 32px));background:#fff;color:#1f2329;border:1px solid #dee0e3;border-radius:12px;box-shadow:0 12px 40px rgba(31,35,41,.18);font:14px/1.5 -apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif}",
      "#feishu-extract-panel header{display:flex;align-items:center;gap:8px;padding:12px 14px;border-bottom:1px solid #eff0f1}",
      "#feishu-extract-panel header strong{flex:1}",
      "#feishu-extract-panel button{border:1px solid #dee0e3;background:#fff;border-radius:8px;padding:6px 10px;cursor:pointer}",
      "#feishu-extract-status{padding:12px 14px;color:#646a73}",
      "#feishu-extract-actions{display:flex;gap:8px;padding:0 14px 14px}",
    ].join("");
    document.documentElement.appendChild(style);

    const button = document.createElement("button");
    button.id = "feishu-extract-btn";
    button.type = "button";
    button.textContent = "提取并打印";
    document.documentElement.appendChild(button);

    const panel = document.createElement("section");
    panel.id = "feishu-extract-panel";
    panel.hidden = true;
    panel.innerHTML = [
      "<header><strong>提取并打印</strong><button type='button' data-act='close'>关闭</button></header>",
      "<div id='feishu-extract-status'>还没有提取</div>",
      "<div id='feishu-extract-actions'>",
      "<button type='button' data-act='open'>打开打印页</button>",
      "<button type='button' data-act='download'>下载 HTML</button>",
      "</div>",
    ].join("");
    document.documentElement.appendChild(panel);

    const status = panel.querySelector("#feishu-extract-status");
    let current = null;

    async function extract() {
      panel.hidden = false;
      current = null;
      status.textContent = "正在读取正文…";
      const target = tokenFromUrl();
      if (!target) {
        status.textContent = "当前网址不是飞书文档或知识库页面。";
        return;
      }
      try {
        const doc = await resolveDoc(target);
        panel.querySelector("header strong").textContent = doc.title;
        const blocks = await fetchBlocks(doc.id);
        const images = collectImages(blocks, doc.id);
        status.textContent = images.length ? "正在下载图片 0/" + images.length + "…" : "这篇没有图片，正在排版…";
        const loaded = await loadImages(images, (done, total) => {
          status.textContent = "正在下载图片 " + done + "/" + total + "…";
        });
        const failed = images.filter((image) => !loaded[image.id]).length;
        const html = renderHtml(blocks, doc.id, doc.title, loaded);
        current = { title: doc.title, html };
        const opened = openPrint(html);
        status.textContent =
          "已生成打印页。" +
          (failed ? failed + " 张图片没下下来。" : "") +
          (opened ? " 在新页面点「打印」。" : " 浏览器拦截了弹窗，请点「打开打印页」或下载 HTML。");
      } catch (error) {
        status.textContent = error && error.message ? error.message : "提取失败";
      }
    }

    function openPrint(html) {
      const printWindow = window.open("", "_blank");
      if (!printWindow) return false;
      printWindow.document.open();
      printWindow.document.write(html);
      printWindow.document.close();
      return true;
    }

    button.addEventListener("click", extract);
    panel.addEventListener("click", (event) => {
      const action = event.target && event.target.getAttribute("data-act");
      if (action === "close") panel.hidden = true;
      if (!current) return;
      if (action === "open" && !openPrint(current.html)) status.textContent = "弹窗被拦截了，请改用「下载 HTML」，用浏览器打开后再打印。";
      if (action === "download") downloadHtml(safeFilename(current.title), current.html);
    });
  }

  buildPanel();
})();
