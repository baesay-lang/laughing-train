/* 대시보드 로직: data.json을 읽어 타임라인/필터/상세보기를 렌더링 */
(async function () {
  const data = await fetch("data.json").then(r => r.json());
  const state = { category: null, project: null, query: "" };

  const $ = sel => document.querySelector(sel);
  const timeline = $("#timeline");
  const overlay = $("#detail-overlay");

  const KIND_ICON = { md: "📄", html: "🌐", pdf: "📕", image: "🖼️", code: "💻", other: "📎" };

  function fmtSize(n) {
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
    return (n / 1024 / 1024).toFixed(1) + " MB";
  }

  function fileHref(f) {
    if (f.kind === "md") return "view.html?f=" + encodeURIComponent(f.path);
    return f.path;
  }

  function matches(item) {
    if (state.category && item.category !== state.category) return false;
    if (state.project && item.project !== state.project) return false;
    if (state.query) {
      const hay = [item.title, item.description, item.category, item.project,
                   (item.tags || []).join(" ")].join(" ").toLowerCase();
      if (!hay.includes(state.query)) return false;
    }
    return true;
  }

  function renderChips(el, values, key, counts) {
    el.innerHTML = "";
    if (!values.length) {
      el.innerHTML = '<span class="meta-note">아직 없음</span>';
      return;
    }
    const all = document.createElement("button");
    all.className = "chip" + (state[key] === null ? " active" : "");
    all.textContent = "전체";
    all.onclick = () => { state[key] = null; render(); };
    el.appendChild(all);
    for (const v of values) {
      const b = document.createElement("button");
      b.className = "chip" + (state[key] === v ? " active" : "");
      b.innerHTML = escapeHtml(v) + '<span class="count">' + (counts[v] || 0) + "</span>";
      b.onclick = () => { state[key] = state[key] === v ? null : v; render(); };
      el.appendChild(b);
    }
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  function render() {
    const catCounts = {}, projCounts = {};
    for (const it of data.items) {
      catCounts[it.category] = (catCounts[it.category] || 0) + 1;
      if (it.project) projCounts[it.project] = (projCounts[it.project] || 0) + 1;
    }
    renderChips($("#category-chips"), data.categories, "category", catCounts);
    renderChips($("#project-chips"), data.projects, "project", projCounts);

    const visible = data.items.filter(matches);
    $("#stats").textContent = `전체 ${data.items.length}건 중 ${visible.length}건 표시`;
    $("#generated").textContent = "갱신: " + (data.generated_at || "").slice(0, 16).replace("T", " ") + " UTC";

    timeline.innerHTML = "";
    if (!visible.length) {
      timeline.innerHTML = '<p class="empty">조건에 맞는 결과물이 없습니다.</p>';
      return;
    }

    const groups = new Map();
    for (const it of visible) {
      const ym = it.date.slice(0, 7);
      if (!groups.has(ym)) groups.set(ym, []);
      groups.get(ym).push(it);
    }

    for (const [ym, items] of groups) {
      const g = document.createElement("div");
      g.className = "month-group";
      const [y, m] = ym.split("-");
      g.innerHTML = `<h2 class="month-heading">${y}년 ${Number(m)}월</h2>`;
      const cards = document.createElement("div");
      cards.className = "cards";
      for (const it of items) cards.appendChild(card(it));
      g.appendChild(cards);
      timeline.appendChild(g);
    }
  }

  function card(it) {
    const el = document.createElement("button");
    el.className = "card";
    const badges = [`<span class="badge">${escapeHtml(it.category)}</span>`];
    if (it.project) badges.push(`<span class="badge project">${escapeHtml(it.project)}</span>`);
    if (it.archive && it.archive.length)
      badges.push(`<span class="badge ver">이전 버전 ${it.archive.length}개</span>`);
    el.innerHTML = `
      <div class="card-top"><h3>${escapeHtml(it.title)}</h3>
        <span class="date">${it.date}</span></div>
      <p class="desc">${escapeHtml(it.description || "")}</p>
      <div class="badges">${badges.join("")}</div>`;
    el.onclick = () => openDetail(it);
    return el;
  }

  function fileListHtml(files) {
    return '<ul class="file-list">' + files.map(f => `
      <li><a href="${fileHref(f)}" target="_blank" rel="noopener">
        <span>${KIND_ICON[f.kind] || "📎"}</span>
        <span>${escapeHtml(f.name)}</span>
        <span class="size">${fmtSize(f.size)}</span></a></li>`).join("") + "</ul>";
  }

  function openDetail(it) {
    let html = `
      <h2>${escapeHtml(it.title)}</h2>
      <p class="detail-meta">${it.date} · ${escapeHtml(it.category)}${it.project ? " · " + escapeHtml(it.project) : ""}</p>
      ${it.description ? `<p class="detail-desc">${escapeHtml(it.description)}</p>` : ""}`;

    if (it.external_url)
      html += `<h4>바로가기</h4><ul class="file-list"><li>
        <a href="${escapeHtml(it.external_url)}" target="_blank" rel="noopener">
        <span>🔗</span><span>${escapeHtml(it.external_url)}</span></a></li></ul>`;

    if (it.files.length) html += `<h4>파일</h4>` + fileListHtml(it.files);

    if (it.archive && it.archive.length) {
      html += `<h4>이전 버전</h4>`;
      for (const v of it.archive) {
        html += `<details class="archive-block"><summary>${escapeHtml(v.version)}</summary>${fileListHtml(v.files)}</details>`;
      }
    }

    html += `<div class="links-row">
      <a href="${it.history_url}" target="_blank" rel="noopener">변경 이력 (GitHub) ↗</a>
      <a href="${data.repo_url}/tree/main/works/${encodeURIComponent(it.id)}" target="_blank" rel="noopener">저장소에서 보기 ↗</a>
    </div>`;

    $("#detail-body").innerHTML = html;
    overlay.classList.remove("hidden");
  }

  $("#detail-close").onclick = () => overlay.classList.add("hidden");
  overlay.onclick = e => { if (e.target === overlay) overlay.classList.add("hidden"); };
  document.addEventListener("keydown", e => {
    if (e.key === "Escape") overlay.classList.add("hidden");
  });
  $("#search").addEventListener("input", e => {
    state.query = e.target.value.trim().toLowerCase();
    render();
  });

  render();
})();
