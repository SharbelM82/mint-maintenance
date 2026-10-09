// Live preview of the calculated (formula) fields while the form is edited.
// The server runs the same calculation code it uses when saving.
(function () {
  const form = document.getElementById("wo-form");
  if (!form) return;
  let timer = null;
  let dirty = false;

  async function refresh() {
    try {
      const res = await fetch(form.dataset.calcUrl, { method: "POST", body: new FormData(form) });
      if (!res.ok) return;
      const out = await res.json();
      for (const [key, text] of Object.entries(out)) {
        const el = document.getElementById("calc-" + key);
        if (!el) continue;
        el.textContent = text || "—";
        el.classList.toggle("neg", /^\(\$/.test(text));
        el.classList.toggle("warn", /^Enter|exceeds/.test(text));
      }
    } catch (e) { /* offline: the values are recalculated on save anyway */ }
  }

  form.addEventListener("input", () => {
    dirty = true;
    clearTimeout(timer);
    timer = setTimeout(refresh, 250);
  });
  form.addEventListener("submit", () => { dirty = false; });
  window.addEventListener("beforeunload", (e) => {
    if (dirty) { e.preventDefault(); e.returnValue = ""; }
  });
  refresh();
})();
