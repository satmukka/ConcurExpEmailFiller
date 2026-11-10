(function () {
  // Toggle: if the panel already exists, remove it.
  const existing = document.getElementById("aef-panel");
  if (existing) {
    existing.remove();
    return;
  }

  let emails = [];
  let stop = false;
  let running = false;

  const panel = document.createElement("div");
  panel.id = "aef-panel";
  panel.style.cssText =
    "position:fixed;top:12px;right:12px;z-index:2147483647;width:290px;" +
    "background:linear-gradient(160deg,#fff7fb 0%,#f1f7ff 50%,#f3fff6 100%);" +
    "border:2px solid #7c4dff;border-radius:12px;padding:12px;font:13px Segoe UI,Arial,sans-serif;" +
    "color:#111;box-shadow:0 4px 18px rgba(124,77,255,.4);user-select:none;";
  const btn = (bg) =>
    "padding:5px 8px;cursor:pointer;font-weight:700;color:#fff;border:none;border-radius:6px;background:" + bg + ";";
  const numBox = (c) => "width:70px;padding:2px 4px;border:2px solid " + c + ";border-radius:5px;";
  panel.innerHTML =
    '<div id="aef-drag-handle" style="font-weight:800;font-size:15px;margin:-12px -12px 10px -12px;' +
    'color:#fff;cursor:move;display:flex;align-items:center;gap:6px;padding:9px 12px;' +
    'background:linear-gradient(90deg,#7c4dff,#e040fb,#ff6d00);border-radius:10px 10px 0 0;user-select:none;">' +
    '<span style="font-size:11px;color:#ffe9ff;letter-spacing:1px">⠿</span>' +
    '<span>✉️ Exp Email Filler</span>' +
    '<span style="font-size:10px;color:#ffe9ff;margin-left:auto;font-weight:400">drag to move</span>' +
    '</div>' +
    '<input type="file" id="aef-file" accept=".xlsx,.xls,.csv" style="width:100%;margin-bottom:8px;font-size:12px;color:#6a1b9a">' +
    '<div style="margin-bottom:6px;font-size:12px;color:#00838f;font-weight:600">Max wait for result (ms): ' +
    '<input id="aef-d1" type="number" value="1000" style="' + numBox("#00bcd4") + '"></div>' +
    '<div style="margin-bottom:8px;font-size:12px;color:#c2185b;font-weight:600">Wait after Add (ms): ' +
    '<input id="aef-d2" type="number" value="200" style="' + numBox("#ec407a") + '"></div>' +
    '<div style="display:flex;gap:6px;margin-bottom:8px;">' +
    '<button id="aef-start" disabled style="flex:1;' + btn("linear-gradient(90deg,#00c853,#64dd17)") + '">Start</button>' +
    '<button id="aef-stop" disabled style="flex:1;' + btn("linear-gradient(90deg,#ff6d00,#ff1744)") + '">Stop</button>' +
    '<button id="aef-close" style="' + btn("linear-gradient(90deg,#5c6bc0,#7e57c2)") + '">Close</button>' +
    '</div>' +
    '<div id="aef-status" style="margin-top:6px;color:#4527a0;font-weight:600;font-size:12px;background:#ede7f6;padding:6px;border-radius:6px;border:1px solid #b39ddb;word-break:break-word">Pick the Excel file.</div>';
  document.body.appendChild(panel);

  const $ = (id) => panel.querySelector("#" + id);
  const status = (t) => {
    $("aef-status").textContent = t;
  };
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const LOG = "[Exp Email Filler]";

  // ── Drag to move panel ────────────────────────────────────────────────────
  (function makeDraggable() {
    const handle = $("aef-drag-handle");
    let dragging = false, startX, startY, origLeft, origTop;

    handle.addEventListener("mousedown", (e) => {
      if (e.button !== 0) return;
      dragging = true;
      startX = e.clientX;
      startY = e.clientY;
      const rect = panel.getBoundingClientRect();
      origLeft = rect.left;
      origTop  = rect.top;
      panel.style.right  = "auto";
      panel.style.left   = origLeft + "px";
      panel.style.top    = origTop  + "px";
      e.preventDefault();
    });

    document.addEventListener("mousemove", (e) => {
      if (!dragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      const newLeft = Math.max(0, Math.min(window.innerWidth - panel.offsetWidth, origLeft + dx));
      const newTop  = Math.max(0, Math.min(window.innerHeight - panel.offsetHeight, origTop + dy));
      panel.style.left = newLeft + "px";
      panel.style.top  = newTop  + "px";
    });

    document.addEventListener("mouseup", () => {
      dragging = false;
    });
  })();

  $("aef-close").onclick = () => {
    stop = true;
    panel.remove();
  };
  $("aef-stop").onclick = () => {
    stop = true;
  };

  $("aef-file").onchange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    try {
      const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
      const isEmail = (cell) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(cell == null ? "" : cell).trim());
      // Use the leftmost column that contains at least one email address.
      const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
      let col = -1;
      for (let c = 0; c < width && col < 0; c++) {
        if (rows.some((r) => isEmail(r[c]))) col = c;
      }
      const seen = new Set();
      emails = [];
      if (col >= 0) {
        for (const row of rows) {
          // Non-email cells in that column (e.g. the header) are skipped.
          if (!isEmail(row[col])) continue;
          const v = String(row[col]).trim();
          if (!seen.has(v.toLowerCase())) {
            seen.add(v.toLowerCase());
            emails.push(v);
          }
        }
      }
      // Column letter for display (0 -> A, 25 -> Z, 26 -> AA).
      let letter = "";
      for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) letter = String.fromCharCode(65 + ((n - 1) % 26)) + letter;
      status(
        col < 0
          ? "No email addresses found in the first sheet."
          : `${emails.length} emails loaded from column ${letter}. Open the search page, then press Start.`
      );
      $("aef-start").disabled = emails.length === 0;
    } catch (err) {
      status("Could not read file: " + err.message);
    }
  };

  // ── Page helpers ──────────────────────────────────────────────────────────
  const isVisible = (el) => {
    if (!el || panel.contains(el)) return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0 && getComputedStyle(el).visibility !== "hidden";
  };

  // Text that labels a form control: <label for>, wrapping <label>, aria-label, placeholder, name, id.
  function labelText(el) {
    const parts = [el.getAttribute("aria-label"), el.placeholder, el.name, el.id];
    if (el.id) {
      const l = document.querySelector('label[for="' + CSS.escape(el.id) + '"]');
      if (l) parts.push(l.textContent);
    }
    const wrap = el.closest("label");
    if (wrap) parts.push(wrap.textContent);
    const by = el.getAttribute("aria-labelledby");
    if (by) by.split(/\s+/).forEach((id) => { const n = document.getElementById(id); if (n) parts.push(n.textContent); });
    return parts.filter(Boolean).join(" ");
  }

  // The "Email Address" text box.
  function findEmailInput() {
    const inputs = [...document.querySelectorAll("input, textarea")].filter(
      (el) => isVisible(el) && !/^(checkbox|radio|button|submit|file|hidden)$/i.test(el.type || "")
    );
    return (
      inputs.find((el) => /e-?mail\s*address/i.test(labelText(el))) ||
      inputs.find((el) => /e-?mail/i.test(labelText(el))) ||
      null
    );
  }

  // A visible button whose text is exactly `name` (case-insensitive).
  function findButton(name) {
    const want = name.toLowerCase();
    const els = [...document.querySelectorAll('button, input[type="button"], input[type="submit"], [role="button"], a')];
    return (
      els.find((el) => {
        if (!isVisible(el)) return false;
        const t = (el.tagName === "INPUT" ? el.value : el.textContent || el.getAttribute("aria-label") || "")
          .trim().toLowerCase();
        return t === want;
      }) || null
    );
  }

  const isDisabled = (el) => el.disabled || el.getAttribute("aria-disabled") === "true";

  // Full pointer/mouse sequence at the element's centre (a single click; includes the click event).
  function fullClick(el) {
    const r = el.getBoundingClientRect();
    const o = { bubbles: true, cancelable: true, composed: true, view: window, button: 0,
      clientX: r.left + r.width / 2, clientY: r.top + r.height / 2 };
    for (const t of ["pointerdown", "mousedown", "pointerup", "mouseup", "click"]) {
      el.dispatchEvent(new (t.startsWith("pointer") ? PointerEvent : MouseEvent)(t, o));
    }
  }

  // Overwrite the text box content (works with React/Angular style inputs).
  function setText(el, text) {
    el.focus();
    if (typeof el.select === "function") el.select();
    const proto = el.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    const descriptor = Object.getOwnPropertyDescriptor(proto, "value");
    if (descriptor && descriptor.set) descriptor.set.call(el, text);
    else el.value = text;
    el.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: true, inputType: "insertFromPaste", data: text }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }

  // Concur's result grid is made of divs: every cell of a row shares the same data-row-key, and the
  // row's own Add is <button data-nuiexp="add: <id>" aria-label="Add: <NAME>">. (The header Add,
  // data-nuiexp="sat-btn-search-results-add", is disabled until a checkbox is ticked, so it is not used.)
  // Returns { tr, add } for the row whose text shows this email, or null.
  function findResultRow(email) {
    const needle = email.toLowerCase();
    for (const add of document.querySelectorAll('button[data-nuiexp^="add:"]')) {
      if (!isVisible(add)) continue;
      const cell = add.closest("[data-row-key]");
      let text;
      if (cell) {
        const key = cell.getAttribute("data-row-key");
        text = [...document.querySelectorAll("[data-row-key]")]
          .filter((c) => c.getAttribute("data-row-key") === key)
          .map((c) => c.textContent)
          .join(" ");
      } else {
        let up = add.parentElement;
        for (let d = 0; up && d < 6 && !(up.textContent || "").toLowerCase().includes(needle); d++) up = up.parentElement;
        text = up ? up.textContent : "";
      }
      if ((text || "").toLowerCase().includes(needle)) return { tr: cell || add.parentElement, add };
    }
    return null;
  }

  async function waitForResultRow(email, maxWait) {
    const start = Date.now();
    while (Date.now() - start < maxWait && !stop) {
      const hit = findResultRow(email);
      if (hit) {
        await sleep(300); // let the table settle
        return findResultRow(email) || hit;
      }
      await sleep(150);
    }
    return null;
  }

  async function waitEnabled(el, maxWait) {
    const start = Date.now();
    while (Date.now() - start < maxWait && !stop) {
      if (!isDisabled(el)) return true;
      await sleep(150);
    }
    return !isDisabled(el);
  }

  $("aef-start").onclick = async () => {
    if (running) return;
    const input = findEmailInput();
    const searchBtn = findButton("Search");
    if (!input || !searchBtn) {
      status(
        "Could not find the " +
          [!input && "'Email Address' box", !searchBtn && "'Search' button"].filter(Boolean).join(" and ") +
          " on this page."
      );
      return;
    }
    running = true;
    stop = false;
    $("aef-start").disabled = true;
    $("aef-stop").disabled = false;
    const d1 = parseInt($("aef-d1").value, 10) || 1000;
    const d2 = parseInt($("aef-d2").value, 10) || 200;

    const added = [];
    const problems = [];
    let i = 0;
    for (; i < emails.length && !stop; i++) {
      const email = emails[i];
      status(`Processing ${i + 1}/${emails.length}: ${email}`);

      // Re-find controls each time in case the page re-rendered.
      const box = findEmailInput();
      const sBtn = findButton("Search");
      if (!box || !sBtn) {
        status(`Page changed before ${email}: 'Email Address' box or 'Search' button is gone.`);
        break;
      }

      // 1. Overwrite the email address and click Search
      setText(box, email);
      await sleep(150);
      sBtn.click();

      // 2. Wait for the result row showing this email
      const hit = await waitForResultRow(email, d1);
      if (!hit) {
        if (stop) break;
        problems.push(email + " (no result)");
        continue;
      }

      // 3. Click the row's own "Add" (no checkbox needed)
      const addBtn = hit.add;
      if (!addBtn) {
        problems.push(email + " ('Add' not found in result row)");
        console.warn(LOG, "No Add control in result row. Row HTML:", hit.tr.outerHTML.slice(0, 1200),
          "| iframes on page:", document.querySelectorAll("iframe").length);
        continue;
      }
      if (!(await waitEnabled(addBtn, 3000))) {
        problems.push(email + " ('Add' button stayed disabled)");
        console.warn(LOG, "Add button disabled:", addBtn.outerHTML.slice(0, 300));
        continue;
      }
      console.log(LOG, "clicking Add:", addBtn.outerHTML.slice(0, 300), "| row:", hit.tr.outerHTML.slice(0, 600));
      addBtn.scrollIntoView({ block: "nearest" });
      await sleep(100);
      fullClick(addBtn);
      added.push(email);
      await sleep(d2);
    }

    const summary = `Added ${added.length}, problems ${problems.length}` +
      (problems.length ? ": " + problems.join(", ") : "");
    console.log(LOG, summary);
    status((stop ? `Stopped at ${i}/${emails.length}. ` : "Done. ") + summary);
    running = false;
    $("aef-start").disabled = false;
    $("aef-stop").disabled = true;
  };
})();
