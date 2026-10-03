const API = "https://app.asana.com/api/1.0";
const LINK_HEADER = "Asana Link";
const $ = (id) => document.getElementById(id);

let current = null; // { rowIndex, headers, values, colCount, linkCol }

function setStatus(msg, ok) {
  const s = $("status");
  s.textContent = msg;
  s.className = ok === undefined ? "" : ok ? "ok" : "err";
}

// Token storage: workbook (travels with the file) or this device only
function getToken() {
  return Office.context.document.settings.get("asanaToken") || localStorage.getItem("asanaToken") || "";
}

function saveToken(tok, inWorkbook) {
  const st = Office.context.document.settings;
  if (inWorkbook) {
    st.set("asanaToken", tok);
    localStorage.removeItem("asanaToken");
  } else {
    localStorage.setItem("asanaToken", tok);
    st.remove("asanaToken");
  }
  st.saveAsync(() => {});
}

function clearToken() {
  const st = Office.context.document.settings;
  st.remove("asanaToken");
  st.saveAsync(() => {});
  localStorage.removeItem("asanaToken");
}

async function api(path, opts = {}) {
  const res = await fetch(API + path, {
    ...opts,
    headers: {
      Authorization: "Bearer " + getToken(),
      "Content-Type": "application/json",
      Accept: "application/json",
    },
  });
  const json = await res.json();
  if (!res.ok) throw new Error(json.errors?.[0]?.message || "HTTP " + res.status);
  return json;
}

// Fetch all pages of a list endpoint
async function apiList(path) {
  let out = [];
  let offset = "";
  do {
    const sep = path.includes("?") ? "&" : "?";
    const json = await api(`${path}${sep}limit=100${offset ? "&offset=" + offset : ""}`);
    out = out.concat(json.data);
    offset = json.next_page?.offset || "";
  } while (offset);
  return out;
}

function fill(select, items, withEmpty) {
  select.innerHTML = "";
  if (withEmpty) select.add(new Option("(no section)", ""));
  items.forEach((i) => select.add(new Option(i.name, i.gid)));
}

async function loadWorkspaces() {
  try {
    setStatus("Loading workspaces...");
    const ws = await apiList("/workspaces");
    fill($("workspace"), ws);
    const saved = localStorage.getItem("asanaWorkspace");
    if (saved) $("workspace").value = saved;
    await loadProjects();
    setStatus("");
  } catch (e) {
    setStatus("Asana error: " + e.message, false);
  }
}

async function loadProjects() {
  const ws = $("workspace").value;
  localStorage.setItem("asanaWorkspace", ws);
  const projects = await apiList(`/projects?workspace=${ws}&archived=false&opt_fields=name`);
  projects.sort((a, b) => a.name.localeCompare(b.name));
  fill($("project"), projects);
  const saved = localStorage.getItem("asanaProject");
  if (saved && projects.some((p) => p.gid === saved)) $("project").value = saved;
  await loadSections();
}

async function loadSections() {
  const p = $("project").value;
  if (!p) return fill($("section"), [], true);
  localStorage.setItem("asanaProject", p);
  const sections = await apiList(`/projects/${p}/sections?opt_fields=name`);
  fill($("section"), sections, true);
}

async function readRow() {
  try {
    await Excel.run(async (ctx) => {
      const sel = ctx.workbook.getSelectedRange();
      sel.load("rowIndex");
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      const used = ws.getUsedRange();
      used.load("columnCount");
      await ctx.sync();

      if (sel.rowIndex === 0) return; // header row
      const n = used.columnCount;
      const hdr = ws.getRangeByIndexes(0, 0, 1, n);
      const row = ws.getRangeByIndexes(sel.rowIndex, 0, 1, n);
      hdr.load("values");
      row.load("values");
      await ctx.sync();

      const headers = hdr.values[0].map(String);
      const values = row.values[0];
      let linkCol = headers.indexOf(LINK_HEADER);
      const hasLinkCol = linkCol >= 0;
      if (!hasLinkCol) linkCol = n;
      current = { rowIndex: sel.rowIndex, headers, values, colCount: n, linkCol, hasLinkCol };

      const prev = $("nameCol").value;
      $("nameCol").innerHTML = "";
      headers.forEach((h, i) => {
        if (h !== LINK_HEADER) $("nameCol").add(new Option(h || "Column " + (i + 1), i));
      });
      $("nameCol").value = prev || "0";
      refreshText();
    });
  } catch (e) {
    setStatus("Excel error: " + e.message, false);
  }
}

function refreshText() {
  if (!current) return;
  const ni = Number($("nameCol").value || 0);
  $("name").value = String(current.values[ni] ?? "");
  const notes = current.headers
    .map((h, i) => ({ h, v: current.values[i], i }))
    .filter((x) => x.i !== ni && x.h !== LINK_HEADER && x.v !== "" && x.v != null)
    .map((x) => `${x.h}: ${x.v}`)
    .join("\n");
  $("notes").value = notes;
}

async function createTask() {
  if (!current) return setStatus("Select a data row first.", false);
  if (!$("name").value.trim()) return setStatus("Task name is empty.", false);
  if (!$("project").value) return setStatus("Choose a project.", false);

  try {
    // Duplicate check: link already in this row
    const existing = current.hasLinkCol ? current.values[current.linkCol] : "";
    if (existing && !confirm("This row already has an Asana link. Create another task?")) return;

    setStatus("Creating task...");
    const data = {
      name: $("name").value.trim(),
      notes: $("notes").value,
      projects: [$("project").value],
    };
    if ($("due").value) data.due_on = $("due").value;
    if ($("section").value) {
      data.memberships = [{ project: $("project").value, section: $("section").value }];
      delete data.projects;
    }

    const res = await api("/tasks?opt_fields=permalink_url", {
      method: "POST",
      body: JSON.stringify({ data }),
    });
    const url = res.data.permalink_url;

    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      if (!current.hasLinkCol) ws.getCell(0, current.linkCol).values = [[LINK_HEADER]];
      ws.getCell(current.rowIndex, current.linkCol).values = [[url]];
      await ctx.sync();
    });

    setStatus("Task created:\n" + url, true);
    await readRow();
  } catch (e) {
    setStatus("Failed: " + e.message, false);
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  $("saveToken").onclick = () => {
    const tok = $("token").value.trim();
    if (!tok) return setStatus("Paste a token first.", false);
    saveToken(tok, $("inWorkbook").checked);
    $("token").value = "";
    $("settings").open = false;
    loadWorkspaces();
  };
  $("clearToken").onclick = () => {
    clearToken();
    $("workspace").innerHTML = "";
    $("project").innerHTML = "";
    $("section").innerHTML = "";
    $("settings").open = true;
    setStatus("Token removed.", true);
  };
  $("workspace").onchange = () => loadProjects().catch((e) => setStatus(e.message, false));
  $("project").onchange = () => loadSections().catch((e) => setStatus(e.message, false));
  $("nameCol").onchange = refreshText;
  $("create").onclick = createTask;

  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, readRow);

  if (getToken()) {
    $("settings").open = false;
    loadWorkspaces();
  } else {
    $("settings").open = true;
  }
  readRow();
});
