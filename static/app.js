/* 经营数据管理系统前端：零构建，原生 JS + 仪表盘样式 */
"use strict";

const state = { meta: null, preview: null };

async function api(url, opts) {
  const res = await fetch(url, opts);
  if (!res.ok) {
    let msg = res.statusText;
    try { msg = (await res.json()).detail || msg; } catch (e) {}
    throw new Error(msg);
  }
  return res.json();
}

// ---------- Toast ----------
function toast(msg, type = "info") {
  const wrap = document.getElementById("toast-wrap");
  const el = document.createElement("div");
  el.className = "toast " + type;
  el.textContent = msg;
  wrap.appendChild(el);
  setTimeout(() => { el.style.opacity = "0"; el.style.transition = "opacity .3s"; }, 2400);
  setTimeout(() => el.remove(), 2800);
}

// ---------- 页签 ----------
const PAGE_META = {
  entry: ["数据录入", "选择产品与月份，直接修改单元格保存数据"],
  import: ["数据导入", "从原始表格解析数据，预览后确认入库"],
  details: ["数据明细", "查看库里所有录入的数据，筛选后一键导出"],
  report: ["报表", "选择模板与期间，渲染预览或导出 Excel"],
};
document.querySelectorAll(".sidebar nav button").forEach((b) => {
  b.onclick = () => {
    document.querySelectorAll(".sidebar nav button").forEach((x) => x.classList.toggle("active", x === b));
    document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("active", x.id === "tab-" + b.dataset.tab));
    const [title, desc] = PAGE_META[b.dataset.tab];
    document.getElementById("page-title").textContent = title;
    document.getElementById("page-desc").textContent = desc;
  };
});

// ---------- 元数据与控件 ----------
function fillSelect(sel, items, text = (x) => x.name, value = (x) => x.id) {
  sel.innerHTML = "";
  for (const it of items) {
    const o = document.createElement("option");
    o.value = value(it);
    o.textContent = text(it);
    sel.appendChild(o);
  }
}

function fillButtonGroup(container, items, opts = {}) {
  /* items: [{id, label}]，点击切换 active；返回当前值读取函数 */
  container.innerHTML = "";
  let current = opts.default;
  const get = () => current;
  for (const it of items) {
    const b = document.createElement("button");
    b.textContent = it.label;
    b.classList.toggle("active", it.id === current);
    b.onclick = () => {
      current = it.id;
      container.querySelectorAll("button").forEach((x) => x.classList.toggle("active", x === b));
      opts.onChange && opts.onChange(it.id);
    };
    container.appendChild(b);
  }
  return get;
}

async function loadMeta() {
  state.meta = await api("/api/meta");
  const y = new Date().getFullYear();
  fillSelect(document.getElementById("entry-year"), [y - 1, y].map((x) => ({ id: x, name: x + "年" })));
  fillSelect(document.getElementById("report-year"), [y - 1, y].map((x) => ({ id: x, name: x + "年" })));
  const months = Array.from({ length: 12 }, (_, i) => ({ id: i + 1, label: i + 1 + "月" }));
  state.getEntryMonth = fillButtonGroup(document.getElementById("entry-month"), months, { default: 7 });
  state.getReportMonth = fillButtonGroup(document.getElementById("report-month"), months, { default: 7 });
  fillSelect(document.getElementById("entry-product"), state.meta.products);
  fillSelect(document.getElementById("report-tpl"), state.meta.report_templates);
  const dtypes = state.meta.data_types.filter((d) => d.base_dtype_id === null);
  state.getEntryDtype = fillButtonGroup(document.getElementById("entry-dtype"),
    dtypes.map((d) => ({ id: d.id, label: d.name })), { default: 1 });
  fillSelect(document.getElementById("import-tpl"), state.meta.source_templates);
  document.getElementById("import-tpl").dispatchEvent(new Event("change"));
  // 数据明细筛选器：年份/月份单选；产品/渠道/数据类型/字段多选（空 = 全部）
  const all = [{ id: "", name: "全部" }];
  fillSelect(document.getElementById("det-year"), all.concat([y - 1, y].map((x) => ({ id: x, name: x + "年" }))));
  for (const id of ["det-month-from", "det-month-to"])
    fillSelect(document.getElementById(id), all.concat(months.map((m) => ({ id: m.id, name: m.label }))));
  const pageReset = () => { detailsState.page = 1; };
  detCtl.product = multiSelect(document.getElementById("det-product"), state.meta.products, () => {
    fillChannelSelect();  // 产品多选联动渠道选项
    pageReset();
  });
  detCtl.channel = multiSelect(document.getElementById("det-channel"), [], pageReset);
  detCtl.dtype = multiSelect(document.getElementById("det-dtype"),
    state.meta.data_types.filter((d) => d.base_dtype_id === null), pageReset);
  detCtl.field = multiSelect(document.getElementById("det-field"), state.meta.fields, pageReset);
  fillChannelSelect();
}

// ---------- 多选下拉（明细筛选器：产品/渠道/数据类型/字段） ----------
// 点击展开选项列表、点击选项切换选中（可多选）、点击外部关闭；空选中 = 「全部」
// 返回 { get, set, rebuild }：get() 返回选中 id 的 Set（空 = 全部）
function multiSelect(root, items, onChange) {
  const sel = new Set();
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "ms-btn";
  btn.innerHTML = `<span class="ms-label"></span><span class="ms-arrow">▾</span>`;
  const panel = document.createElement("div");
  panel.className = "ms-panel";
  root.appendChild(btn);
  root.appendChild(panel);

  const nameOf = (id) => (items.find((it) => String(it.id) === String(id)) || {}).name || id;
  const render = () => {
    btn.querySelector(".ms-label").textContent =
      sel.size === 0 ? "全部" : sel.size === 1 ? nameOf([...sel][0]) : `已选 ${sel.size} 项`;
    let h = `<div class="ms-head"><button type="button" class="ms-link" data-a="all">全选</button>
             <button type="button" class="ms-link" data-a="none">清空</button></div>`;
    for (const it of items) {
      const on = sel.has(String(it.id));
      h += `<div class="ms-item${on ? " on" : ""}" data-id="${it.id}">
              <span class="ms-tick">${on ? "✓" : ""}</span>${it.name}</div>`;
    }
    panel.innerHTML = h;
    panel.querySelectorAll(".ms-item").forEach((el) => {
      el.onclick = () => {
        const id = el.dataset.id;
        sel.has(id) ? sel.delete(id) : sel.add(id);
        render();
        onChange && onChange();
      };
    });
    panel.querySelectorAll("[data-a]").forEach((el) => {
      el.onclick = () => {
        sel.clear();
        if (el.dataset.a === "all") items.forEach((it) => sel.add(String(it.id)));
        render();
        onChange && onChange();
      };
    });
  };
  btn.onclick = () => panel.classList.toggle("open");
  document.addEventListener("click", (e) => {  // 点击外部关闭
    if (!root.contains(e.target)) panel.classList.remove("open");
  });
  render();
  return {
    get: () => sel,
    set: (ids) => { sel.clear(); ids.forEach((id) => sel.add(String(id))); render(); },
    rebuild: (newItems) => {
      items = newItems;
      for (const id of [...sel]) if (!items.some((it) => String(it.id) === String(id))) sel.delete(id);
      render();
    },
  };
}

// 渠道下拉：按产品多选联动；节点显示完整路径（聚合节点 = 该子树全部叶子）
const detCtl = { product: null, channel: null, dtype: null, field: null };
const prodIdMap = () => Object.fromEntries(state.meta.products.map((p) => [p.name, String(p.id)]));
function fillChannelSelect() {
  const prods = detCtl.product.get();
  const prodIdByName = prodIdMap();
  const nodes = state.meta.channel_trees.filter(
    (n) => !prods.size || prods.has(prodIdByName[n.product]));
  const byId = Object.fromEntries(state.meta.channel_trees.map((n) => [n.id, n]));
  const pathOf = (id) => {
    const p = [];
    let cur = id;
    while (cur) { p.unshift(byId[cur].name); cur = byId[cur].parent_id; }
    return p.join("-");
  };
  const items = nodes.map((n) => ({
    id: n.id,
    name: (byId[n.parent_id] ? "" : n.product + " / ") + pathOf(n.id),
  }));
  detCtl.channel.rebuild(items);
}

// ---------- 数据明细 ----------
const detailsState = { page: 1, sort: "ym", dir: "asc" };

// 明细行 HTML（数据明细页 / 钻取弹窗共用）——库内字符串（渠道名/来源）来自导入工作簿，一律转义
function detailRowHtml(d, extraCls = "") {
  const calc = d.source === "计算值";
  return `<tr class="${extraCls}"><td>${d.year}</td><td>${d.month}月</td><td>${esc(d.product)}</td>
    <td style="text-align:left">${esc(d.channel)}</td><td>${esc(d.field)}</td><td>${esc(d.dtype)}</td><td>${esc(fmt(d.value))}</td>
    <td>${calc ? '<span class="badge warn">计算值</span>' : esc(d.source || "")}</td></tr>`;
}

function detailsParams() {
  const p = new URLSearchParams();
  const year = document.getElementById("det-year").value;
  if (year !== "") p.set("year", year);
  const mf = document.getElementById("det-month-from").value, mt = document.getElementById("det-month-to").value;
  if (mf !== "") p.set("month_from", mf);
  if (mt !== "") p.set("month_to", mt);
  // 多选维度（空 = 全部，不传参数）
  for (const [key, ctl] of [["product_ids", detCtl.product], ["channel_ids", detCtl.channel],
                            ["dtype_ids", detCtl.dtype], ["field_ids", detCtl.field]]) {
    const ids = ctl.get();
    if (ids.size) p.set(key, [...ids].join(","));
  }
  if (document.getElementById("det-formula").checked) p.set("include_formula", 1);
  p.set("page", detailsState.page);
  p.set("page_size", 500);
  p.set("sort", detailsState.sort);
  p.set("sort_dir", detailsState.dir);
  return p;
}

async function loadDetails() {
  const btn = document.getElementById("det-load");
  btn.disabled = true;
  try {
    const r = await api("/api/details?" + detailsParams().toString());
    const el = document.getElementById("det-grid");
    const sorted = (k) => detailsState.sort === k ? (detailsState.dir === "asc" ? " ↑" : " ↓") : "";
    let h = `<table><thead><tr>
      <th data-sort="ym">年${sorted("ym")}</th><th data-sort="ym">月</th>
      <th data-sort="product">产品${sorted("product")}</th><th data-sort="channel">渠道${sorted("channel")}</th>
      <th data-sort="field">字段${sorted("field")}</th><th data-sort="dtype">数据类型${sorted("dtype")}</th>
      <th data-sort="value">值${sorted("value")}</th><th>来源</th></tr></thead><tbody>`;
    for (const row of r.rows) h += detailRowHtml(row);
    h += "</tbody></table>";
    el.innerHTML = r.rows.length ? h : `<div class="empty-hint">没有符合条件的数据</div>`;
    document.getElementById("det-count").textContent = `共 ${r.total} 条`;
    document.getElementById("det-count").style.display = "inline-block";
    document.getElementById("det-status").textContent =
      `第 ${r.page} 页 / 共 ${Math.ceil(r.total / r.page_size)} 页`;
    // 分页
    const totalPages = Math.max(1, Math.ceil(r.total / r.page_size));
    let pg = `<button class="btn sm ghost" ${detailsState.page <= 1 ? "disabled" : ""} data-pg="${detailsState.page - 1}">上一页</button>`;
    pg += ` <span style="margin:0 8px">${detailsState.page} / ${totalPages}</span>`;
    pg += `<button class="btn sm ghost" ${detailsState.page >= totalPages ? "disabled" : ""} data-pg="${detailsState.page + 1}">下一页</button>`;
    document.getElementById("det-pager").innerHTML = pg;
    document.getElementById("det-pager").querySelectorAll("button[data-pg]").forEach((b) => {
      b.onclick = () => { detailsState.page = +b.dataset.pg; loadDetails(); };
    });
    el.querySelectorAll("th[data-sort]").forEach((th) => {
      th.style.cursor = "pointer";
      th.onclick = () => {
        const k = th.dataset.sort;
        if (detailsState.sort === k) detailsState.dir = detailsState.dir === "asc" ? "desc" : "asc";
        else { detailsState.sort = k; detailsState.dir = "asc"; }
        detailsState.page = 1;
        loadDetails();
      };
    });
    return r;
  } catch (e) {
    toast("查询失败: " + e.message, "err");
  } finally {
    btn.disabled = false;
  }
}
document.getElementById("det-load").onclick = () => { detailsState.page = 1; loadDetails(); };
document.getElementById("det-formula").onchange = () => { detailsState.page = 1; loadDetails(); };
document.getElementById("det-month-from").onchange = () => { detailsState.page = 1; };
document.getElementById("det-month-to").onchange = () => { detailsState.page = 1; };

document.getElementById("det-export").onclick = () => {
  const p = detailsParams();
  p.delete("page"); p.delete("page_size"); p.delete("sort"); p.delete("dir");
  window.location = "/api/details/export?" + p.toString();
  toast("正在导出明细数据…", "ok");
};
document.getElementById("import-tpl").onchange = (e) => {
  const t = state.meta.source_templates.find((x) => x.id == e.target.value);
  document.getElementById("import-path").value = t ? t.path_pattern : "";
};

// ---------- 数据录入 ----------
let entryLoadSeq = 0;  // 录入网格加载序号（时序守卫）

async function loadGrid() {
  const btn = document.getElementById("entry-load");
  btn.disabled = true;
  btn.textContent = "加载中…";
  const seq = ++entryLoadSeq;  // 请求序号：快速切换加载时丢弃过期响应
  try {
    const year = +document.getElementById("entry-year").value;
    const month = state.getEntryMonth();
    const productId = +document.getElementById("entry-product").value;
    const dtypeId = state.getEntryDtype();
    const facts = await api(`/api/facts?year=${year}&month=${month}&product_id=${productId}`);
    if (seq !== entryLoadSeq) return;  // 已有更新的请求，丢弃过期结果（网格与下拉保持一致）
    const dtype = state.meta.data_types.find((d) => d.id === dtypeId);
    const prodName = state.meta.products.find((p) => p.id === productId).name;

    const nodes = state.meta.channel_trees.filter((n) => n.product === prodName);
    const leaves = nodes.filter((n) => n.is_leaf === 1);
    const nodeById = Object.fromEntries(nodes.map((n) => [n.id, n]));
    const pathOf = (id) => {
      const p = [];
      let cur = id;
      while (cur) { p.unshift(nodeById[cur].name); cur = nodeById[cur].parent_id; }
      return p.join(" / ");
    };

    // 公式字段由引擎计算（不落库），录入网格只显示可编辑的 raw 字段
    const fields = state.meta.fields.filter((f) => f.kind !== "formula");
    const html = ["<table><thead><tr><th>渠道</th>"];
    for (const f of fields) html.push(`<th>${f.name}${f.unit ? "(" + f.unit + ")" : ""}</th>`);
    html.push("</tr></thead><tbody>");
    for (const l of leaves) {
      html.push(`<tr><th>${pathOf(l.id)}</th>`);
      for (const f of fields) {
        const v = facts[l.id] && facts[l.id][f.id] && facts[l.id][f.id][dtypeId];
        html.push(
          `<td><input class="${v === undefined ? "empty" : ""}" data-node="${l.id}" data-field="${f.id}" ` +
          `value="${v === undefined ? "" : v}"></td>`
        );
      }
      html.push("</tr>");
    }
    html.push("</tbody></table>");
    const wrap = document.getElementById("entry-grid");
    wrap.innerHTML = html.join("");
    wrap.querySelectorAll("input").forEach((inp) => {
      inp.onchange = async () => {
        const v = inp.value.trim();
        const num = v === "" ? null : +v;
        if (num !== null && !isFinite(num)) {
          toast("请输入数字（清空格子=删除数据）", "err");
          return;
        }
        try {
          await api("/api/fact", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              year, month, dtype_id: dtypeId,
              node_id: +inp.dataset.node, field_id: +inp.dataset.field,
              value: num,
            }),
          });
          inp.classList.toggle("empty", v === "");
          inp.classList.remove("saved");
          void inp.offsetWidth; // 重启动画
          inp.classList.add("saved");
          document.getElementById("entry-status").textContent =
            `已保存：${year}年${month}月 ${prodName} ${dtype.name} · ${new Date().toLocaleTimeString()}`;
        } catch (e) {
          toast("保存失败: " + e.message, "err");
        }
      };
    });
    document.getElementById("entry-status").textContent =
      `已加载 ${year}年${month}月 ${prodName} ${dtype.name}（${leaves.length} 个渠道 × ${fields.length} 个字段）`;
  } catch (e) {
    toast("加载失败: " + e.message, "err");
  } finally {
    btn.disabled = false;
    btn.textContent = "加载数据";
  }
}
document.getElementById("entry-load").onclick = loadGrid;

// ---------- 导入 ----------
function fmt(x) {
  return x === null || x === undefined ? "" : (typeof x === "number" ? Math.round(x * 100) / 100 : x);
}
// HTML 转义（导入预览内容来自外部工作簿，防止内嵌脚本执行）
function esc(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function fmtPct(x) {
  return x === null || x === undefined ? "" : (x * 100).toFixed(1) + "%";
}
// 趋势着色：同比/达成率类口径 变好达标=绿、变坏未达标=红（仅白名单口径，普通数值列不着色）
// BP达成率 = 比率（>=100% 达标）；其余为差值语义（>=0 变好；净利同比 正差恒为减亏/增盈=好）
const TREND_MEASURES = new Set(["BP达成率", "同比", "毛利率同比", "净利同比", "ASP达成率", "ASP同比"]);
function trendClass(m, v) {
  if (!TREND_MEASURES.has(m) || v === null || v === undefined) return "";
  return v >= (m === "BP达成率" ? 1 : 0) ? "good" : "bad";
}

// 半数远离零取整（与后端 _round_away / Excel ROUND 语义一致；JS Math.round 对负 .5 向 +∞ 会差 1）
function roundAway(x, nd = 0) {
  const mul = 10 ** nd, y = x * mul;
  return (y >= 0 ? Math.floor(y + 0.5) : Math.ceil(y - 0.5)) / mul;
}
// 原表文本口径：毛利率同比 = 个点差（0.032 -> 增加3.2个点）；净利同比 = 盈亏万差（四态）
function fmtMarginYoy(v) {
  const d = Math.abs(roundAway(v * 100, 1));
  return v >= 0 ? `增加${d}个点` : `减少${d}个点`;
}
function fmtProfitYoy(v, yoy) {
  const d = roundAway(v, 0);
  // 今年亏损（或去年同期未知）：正差=减亏 负差=增亏；今年盈利：正差=增盈 负差=减盈
  let verb;
  if (yoy === null || yoy === undefined || yoy + v < 0) verb = d >= 0 ? "减亏" : "增亏";
  else verb = d >= 0 ? "增盈" : "减盈";
  return `${verb}${Math.abs(d)}万`;
}
// 同行同块（group）的 同期 列值（净利同比 文本判断盈亏状态用）
function yoyValue(row, ci, cols) {
  const c = cols[ci];
  for (let j = 0; j < cols.length; j++) {
    if (j !== ci && cols[j].group === c.group && cols[j].measure === "同期") {
      const v = row.cells[j];
      return v === null || v === undefined ? null : v;
    }
  }
  return null;
}

async function previewImport() {
  const tpl = +document.getElementById("import-tpl").value;
  const path = document.getElementById("import-path").value.trim();
  if (!path) { toast("请填写文件路径", "err"); return; }
  const btn = document.getElementById("import-preview");
  btn.disabled = true;
  try {
    const r = await api("/api/import/preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ template_id: tpl, file_path: path }),
    });
    state.preview = { template_id: tpl, file_path: path };
    const el = document.getElementById("import-result");
    el.style.display = "block";
    let h = `<div class="card-title">预览：${r.template_name}</div>`;
    h += `<div class="summary">解析 <b>${r.summary["条数"]}</b> 条
      <span class="badge ok">${r.warnings_total} 条警告</span>
      <span class="badge ${r.conflicts_total ? "bad" : "ok"}">${r.conflicts_total} 条冲突</span></div>`;
    if (r.summary["产品分布"]) {
      h += `<div class="summary">产品分布：` +
        Object.entries(r.summary["产品分布"]).map(([k, v]) => `${k} ${v}`).join(" · ") + `</div>`;
    }
    for (const w of r.warnings.slice(0, 10)) h += `<div class="warn-line">⚠ ${esc(w)}</div>`;
    if (r.conflicts_total) {
      h += `<div class="bad-line">冲突（入库将覆盖现有值）：</div>`;
      for (const c of r.conflicts.slice(0, 8)) {
        h += `<div class="bad-line">&nbsp;&nbsp;${esc(c.item[0])}-${esc(c.item[1])} 节点${esc(c.item[2])} 字段${esc(c.item[3])} 类型${esc(c.item[4])}: ${esc(fmt(c.item[5]))} → 现有 ${esc(fmt(c.old))}</div>`;
      }
    }
    if (r.items_sample.length) {
      h += `<div class="table-wrap" style="margin-top:10px"><table><tr><th>年</th><th>月</th><th>节点</th><th>字段</th><th>类型</th><th>值</th></tr>`;
      for (const it of r.items_sample) h += `<tr>${it.map((x) => `<td>${esc(fmt(x))}</td>`).join("")}</tr>`;
      h += `</table></div>`;
    }
    el.innerHTML = h;
    document.getElementById("import-commit").disabled = false;
    toast("解析完成，确认无误后入库", "ok");
  } catch (e) {
    const el = document.getElementById("import-result");
    el.style.display = "block";
    el.innerHTML = `<div class="bad-line">预览失败: ${e.message}</div>`;
  } finally {
    btn.disabled = false;
  }
}
document.getElementById("import-preview").onclick = previewImport;

async function commitImport() {
  if (!state.preview) return;
  const btn = document.getElementById("import-commit");
  btn.disabled = true;
  try {
    const r = await api("/api/import/commit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(state.preview),
    });
    const el = document.getElementById("import-result");
    el.insertAdjacentHTML("beforeend",
      `<div class="msg-ok" style="margin-top:8px">✅ 入库完成：新写入 ${r["写入"]} 条，覆盖 ${r["覆盖"]} 条</div>`);
    toast(`入库完成：写入 ${r["写入"]} 条`, "ok");
  } catch (e) {
    toast("入库失败: " + e.message, "err");
    btn.disabled = false;
  }
}
document.getElementById("import-commit").onclick = commitImport;

// ---------- 报表 ----------
let reportLoadSeq = 0;  // 报表渲染序号（时序守卫：快速切换模板时丢弃过期响应）

async function renderReport() {
  const seq = ++reportLoadSeq;
  const tpl = +document.getElementById("report-tpl").value;
  const year = +document.getElementById("report-year").value;
  const month = state.getReportMonth();
  const btn = document.getElementById("report-render");
  btn.disabled = true;
  try {
    const r = await api(`/api/report/render?template_id=${tpl}&year=${year}&month=${month}`);
    if (seq !== reportLoadSeq) return;  // 已有更新的渲染请求
    const el = document.getElementById("report-result");
    el.style.display = "block";
    // 组合并：连续同组列合并表头
    const groups = [];
    for (const c of r.columns) {
      const g = c.group || c.measure;
      if (groups.length && groups[groups.length - 1].name === g) groups[groups.length - 1].n++;
      else groups.push({ name: g, n: 1 });
    }
    let h = `<div class="card-title">${r.title} · ${year}年${month}月（${r.rows.length} 行 × ${r.columns.length} 列）</div>`;
    h += `<div class="table-wrap"><table><thead><tr><th rowspan="2">${r.title}</th><th rowspan="2">项目</th>`;
    for (const g of groups) h += `<th colspan="${g.n}">${g.name}</th>`;
    h += `</tr><tr>`;
    // 表头显示名：渠道在列（费用分析）显示渠道名；差值同比类口径统一显示「同比」（内部口径仍区分）
    for (const c of r.columns) {
      const label = c.channel_name || (c.measure.endsWith("同比") ? "同比" : c.measure);
      h += `<th>${label}</th>`;
    }
    h += "</tr></thead><tbody>";
    for (let ri = 0; ri < r.rows.length; ri++) {
      const row = r.rows[ri];
      let span = 1;
      while (ri + span < r.rows.length && r.rows[ri + span].entity === row.entity) span++;
      const first = ri === 0 || r.rows[ri - 1].entity !== row.entity;
      h += `<tr>${first ? `<th rowspan="${span}">${esc(row.entity)}</th>` : ""}`;
      h += `<th style="font-weight:400;color:var(--text-2)">${esc(row.label || "")}</th>`;
      row.cells.forEach((v, i) => {
        const c = r.columns[i];
        // 行级口径覆盖后的实际口径（预测3 毛利率/净利行的同比 -> 差值口径）
        const m = (row.measure_overrides || {})[i] || c.measure;
        const empty = v === null || v === undefined ? " empty" : "";
        // 显示：毛利率同比=个点差、净利同比=盈亏万差（原表文本）；其余按百分比/数值
        let text = "";
        if (!empty && m === "毛利率同比") text = fmtMarginYoy(v);
        else if (!empty && m === "净利同比") text = fmtProfitYoy(v, yoyValue(row, i, r.columns));
        else if (!empty) {
          // 单位取行字段；旧式模板（月度/累计）行无字段，回退列字段
          const fname = row.field || c.field;
          const isPct = (m === "BP达成率" || m === "同比") ||
                        ((state.meta.fields.find((f) => f.name === fname) || {}).unit === "%");
          text = isPct ? fmtPct(v) : fmt(v);
        }
        // 趋势着色：同比/达成率类 变好达标=绿 变坏未达标=红（行级覆盖后的口径判定）
        const tc = trendClass(m, v);
        // 有值单元格可交互：单击=计算过程（公式指标），双击=钻取明细
        h += `<td class="cell${empty}${tc ? " " + tc : ""}" data-ri="${ri}" data-ci="${i}">${text}</td>`;
      });
      h += "</tr>";
    }
    h += "</tbody></table></div>";
    el.innerHTML = h;
    state.report = { columns: r.columns, rows: r.rows, tpl, year, month };
  } catch (e) {
    const el = document.getElementById("report-result");
    el.style.display = "block";
    el.innerHTML = `<div class="bad-line">渲染失败: ${e.message}</div>`;
  } finally {
    btn.disabled = false;
  }
}
document.getElementById("report-render").onclick = renderReport;

// ---------- 单元格交互：单击=计算过程（防双击误触，220ms 延迟），双击=钻取明细 ----------
let cellClickTimer = null;

function onCellClick(e) {
  const td = e.target.closest("td.cell:not(.empty)");
  if (!td) return;
  clearTimeout(cellClickTimer);
  cellClickTimer = setTimeout(() => openExplain(+td.dataset.ri, +td.dataset.ci), 220);
}

function onCellDblClick(e) {
  const td = e.target.closest("td.cell:not(.empty)");
  if (!td) return;
  clearTimeout(cellClickTimer);
  drillFromReport(+td.dataset.ri, +td.dataset.ci);
}

document.getElementById("report-result").addEventListener("click", onCellClick);
document.getElementById("report-result").addEventListener("dblclick", onCellDblClick);

// ---------- 计算过程弹窗 ----------
async function openExplain(ri, ci) {
  // 用上次渲染时的参数请求（改选下拉未重新渲染时，行索引属于上次渲染结果，两者保持一致）
  const { tpl, year, month } = state.report;
  try {
    const tree = await api(`/api/report/explain?template_id=${tpl}&year=${year}&month=${month}&row=${ri}&col=${ci}`);
    const row = state.report.rows[ri];
    document.getElementById("explain-title").textContent =
      `${row.entity || ""} ${row.label || row.field || ""} · ${year}年${month}月`;
    document.getElementById("explain-body").innerHTML = explainNodeHtml(tree, 0);
    // 根的直接子层默认展开
    document.querySelectorAll("#explain-body > .ex-row > .ex-children").forEach((el) => {
      el.style.display = "block";
      const b = el.previousElementSibling.querySelector(".ex-toggle");
      if (b) b.textContent = "▾";
    });
    document.getElementById("explain-modal").style.display = "flex";
  } catch (err) {
    toast("计算过程获取失败: " + err.message, "err");
  }
}

function closeExplain() {
  document.getElementById("explain-modal").style.display = "none";
}

function toggleExplain(btn) {
  const kids = btn.parentElement.querySelector(":scope > .ex-children");
  if (!kids) return;
  const show = kids.style.display !== "block";
  kids.style.display = show ? "block" : "none";
  btn.textContent = show ? "▾" : "▸";
}

function explainNodeHtml(node, depth) {
  const hasKids = node.children && node.children.length;
  const val = node.value === null || node.value === undefined ? "无数据" : fmt(node.value);
  const pad = `style="padding-left:${depth * 18}px"`;
  let h = `<div class="ex-row" ${pad}>`;
  if (hasKids) {
    h += `<button class="ex-toggle" onclick="toggleExplain(this)">▸</button>`;
    h += `<span class="ex-name">${esc(node.name || "")}</span>`;
    if (node.formula) h += `<span class="ex-formula"> = ${esc(node.formula)}</span>`;
    h += `<span class="ex-value"> = ${val}</span>`;
    h += `<div class="ex-children" style="display:none">` +
         node.children.map((c) => explainNodeHtml(c, depth + 1)).join("") + `</div>`;
  } else if (node.raw) {
    h += `<span class="ex-dot">•</span><span class="ex-name">${esc(node.name)}</span>`;
    h += `<span class="ex-value"> = ${val}</span><span class="ex-note">（录入值，叶子汇总）</span>`;
  } else {
    h += `<span class="ex-name">${esc(node.name)}</span><span class="ex-value"> = ${val}</span>`;
  }
  return h + "</div>";
}

// 同步明细页筛选器到钻取条件（年份/月份/四维多选/包含计算值），钻取弹窗与明细页共用同一组条件
function setDrillFilters(seg, nids, field) {
  const yearSel = document.getElementById("det-year");
  // 钻取年份可能不在下拉选项（如报表 2025 年的同期 → 2024）：动态补入选项
  if (![...yearSel.options].some((o) => +o.value === seg.year)) {
    const op = document.createElement("option");
    op.value = seg.year;
    op.textContent = seg.year + "年";
    yearSel.appendChild(op);
  }
  yearSel.value = seg.year;
  document.getElementById("det-month-from").value = seg.months[0];
  document.getElementById("det-month-to").value = seg.months[1];
  const prodIds = new Set();
  const prodIdByName = prodIdMap();
  for (const n of state.meta.channel_trees) if (nids.includes(n.id)) prodIds.add(prodIdByName[n.product]);
  detCtl.product.set([...prodIds].filter(Boolean));
  fillChannelSelect();  // 渠道选项按产品重建后再勾选
  detCtl.channel.set(nids.map(String));
  detCtl.dtype.set((seg.dtype_ids || []).map(String));  // 空 = 全部
  detCtl.field.set(field ? [String(field.id)] : []);
  document.getElementById("det-formula").checked = !!(field && field.kind === "formula");
  detailsState.page = 1;
}

// ---------- 双击钻取：报表单元格 -> 明细弹窗（复合期间多段逐段查询合并） ----------
async function drillFromReport(ri, ci) {
  const row = state.report.rows[ri], col = state.report.columns[ci];
  let drill = col.drill;
  const nids = row.nids || col.channel_nids;
  if (!drill || !drill.segments || !nids || !nids.length) {
    toast("该单元格无明细可钻取", "warn");
    return;
  }
  closeExplain();  // 慢双击时先关掉单击打开的弹窗
  // 行级口径覆盖（费用分析收入行 BP比 → BP 原值）：按覆盖口径构造钻取条件
  const ov = (row.measure_overrides || {})[ci];
  if (ov) {
    const m = state.meta.measures.find((x) => x.name === ov);
    if (m && m.kind === "dtype") {
      drill = { note: m.name, segments: drill.segments.map((s) => ({ ...s, dtype_ids: [m.dtype_id] })) };
    } else if (m) {
      drill = { note: ov, segments: drill.segments.map((s) => ({ ...s, dtype_ids: [] })) };
    }
  }
  // 旧式模板（月度/累计报表）字段在列上；新式行有字段
  const fname = row.field || col.field;
  const field = state.meta.fields.find((f) => f.name === fname);
  try {
    const all = [];
    let total = 0;
    for (const seg of drill.segments) {
      setDrillFilters(seg, nids, field);
      const r = await api("/api/details?" + detailsParams().toString());
      all.push(...r.rows);
      total += r.total;
    }
    const titleParts = drill.segments.map((s) =>
      `${s.year}年${s.months[0] === s.months[1] ? s.months[0] + "月" : s.months[0] + "-" + s.months[1] + "月"}` +
      (s.dtype_ids && s.dtype_ids.length ? "·" + s.dtype_ids.map((i) =>
        (state.meta.data_types.find((d) => d.id === i) || {}).name || "").join("/") : ""));
    document.getElementById("drill-title").textContent =
      `${row.entity || ""} ${row.label || row.field || ""} · ${drill.note}` +
      (drill.segments.length > 1 ? `（${drill.segments.length} 段：${titleParts.join("；")}）` : "");
    const body = document.getElementById("drill-body");
    if (!all.length) {
      body.innerHTML = `<div class="empty-hint">没有符合条件的数据</div>`;
    } else {
      let h = `<div class="grid-wrap"><table><thead><tr>
        <th>年</th><th>月</th><th>产品</th><th>渠道</th><th>字段</th><th>数据类型</th><th>值</th><th>来源</th>
        </tr></thead><tbody>`;
      all.forEach((d, idx) => { h += detailRowHtml(d, idx === 0 ? "drill-hit" : ""); });
      h += "</tbody></table></div>";
      if (total > all.length) h += `<div class="status-line">仅显示前 ${all.length} 条（共 ${total} 条）</div>`;
      body.innerHTML = h;
    }
    document.getElementById("drill-modal").style.display = "flex";
  } catch (e) {
    toast("钻取明细失败: " + e.message, "err");
  }
}

function closeDrill() {
  document.getElementById("drill-modal").style.display = "none";
}

document.getElementById("report-export").onclick = () => {
  const tpl = +document.getElementById("report-tpl").value;
  const year = +document.getElementById("report-year").value;
  const month = state.getReportMonth();
  window.location = `/api/report/export?template_id=${tpl}&year=${year}&month=${month}`;
  toast("已导出 Excel，请查看下载", "ok");
};

loadMeta();
