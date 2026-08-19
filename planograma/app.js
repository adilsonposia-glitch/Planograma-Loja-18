(() => {
  const DATA = window.PLANOGRAMA_DATA;
  if (!DATA) {
    document.body.innerHTML = "<p style='padding:24px;color:#fff'>Dados não carregados (data.js).</p>";
    return;
  }

  const state = {
    filterTipo: "ALL",
    query: "",
    selectedId: null,
    view: "planta",
  };

  const els = {
    kpis: document.getElementById("kpis"),
    pecasBoard: document.getElementById("pecasBoard"),
    floorMap: document.getElementById("floorMap"),
    catList: document.getElementById("catList"),
    gondolaGrid: document.getElementById("gondolaGrid"),
    detail: document.getElementById("detail"),
    detailBody: document.getElementById("detailBody"),
    filterTipo: document.getElementById("filterTipo"),
    search: document.getElementById("search"),
    storeLabel: document.getElementById("storeLabel"),
    navButtons: [...document.querySelectorAll("[data-nav]")],
    helpModal: document.getElementById("helpModal"),
    btnHelp: document.getElementById("btnHelp"),
    btnCloseHelp: document.getElementById("btnCloseHelp"),
  };

  const typeOrder = ["GO", "GB", "BR", "BA", "BP", "GP", "GS", "MQ", "IC", "MH", "MM", "MS", "EC", "EX", "GC"];

  function statusOf(g) {
    if (g.preenchimento === "completo") return "ok";
    if (g.preenchimento === "parcial") return "warn";
    return "alert";
  }

  function fillScore(g) {
    if (g.totalMods <= 0) return 0;
    if (g.preenchimento === "completo") return 100;
    // partial: has modules declared but no shelf BOM
    return 55;
  }

  function filteredGondolas() {
    const q = state.query.trim().toLowerCase();
    return DATA.gondolas.filter((g) => {
      if (state.filterTipo !== "ALL" && g.tipo !== state.filterTipo) return false;
      if (!q) return true;
      const blob = [
        g.id,
        g.setor,
        g.categorias,
        g.tipoLabel,
        ...(g.faces || []).flatMap((f) => [
          f.categoria,
          ...(f.modulos || []).map((m) => [m.categoria, m.especificacao, m.observacao].join(" ")),
        ]),
      ]
        .join(" ")
        .toLowerCase();
      return blob.includes(q);
    });
  }

  function renderKpis() {
    const s = DATA.summary;
    const t = s.totalizador;
    const c = s.contagem;
    const goPrat = DATA.gondolas.filter((g) => g.tipo === "GO").reduce((a, g) => a + g.totalPrat, 0);
    const items = [
      { label: "Módulos (totalizador)", value: t.modulos, hint: `${t.m13x40} × 1,30×0,40 + outros` },
      { label: "Prateleiras", value: t.prateleiras, hint: `GO detalhadas: ${goPrat}` },
      { label: "Braços", value: t.bracos.toLocaleString("pt-BR"), hint: `${t.aparadores || 0} anteparos · ${t.ganchos} ganchos` },
      {
        label: "Equipamentos",
        value: c.gondolas,
        hint: `${c.go} GO · ${c.br} BR · ${c.ba || 0} BA · ${c.bp} BP · ${c.gp || 0} GP · ${c.ic || 0} IC · ${c.mh || 0} MH · ${c.ec || 0} EC`,
      },
    ];
    els.kpis.innerHTML = items
      .map(
        (i) => `<article class="kpi glass">
        <div class="label">${i.label}</div>
        <div class="value">${i.value}</div>
        <div class="hint">${i.hint}</div>
      </article>`
      )
      .join("");
    els.storeLabel.textContent = `${s.loja} · ${s.pavimento}`;
  }

  function renderPecas() {
    const pecas = DATA.summary.pecas || [];
    if (!els.pecasBoard) return;

    const grupos = [];
    pecas.forEach((p) => {
      let g = grupos.find((x) => x.nome === p.grupo);
      if (!g) {
        g = { nome: p.grupo, itens: [], total: 0 };
        grupos.push(g);
      }
      g.itens.push(p);
      g.total += Number(p.qtd) || 0;
    });

    const t = DATA.summary.totalizador || {};
    const resumoRapido = [
      { txt: `${t.modulos || 0} módulos`, sub: "todas as medidas" },
      { txt: `${t.prateleiras || 0} prateleiras`, sub: "todas as medidas" },
      { txt: `${(t.bracos || 0).toLocaleString("pt-BR")} braços`, sub: "35 + 40 + 60 cm" },
      { txt: `${t.aparadores || 0} anteparos`, sub: `aramado · ${t.anteparoAcrilico || 0} acrílico` },
    ];

    els.pecasBoard.innerHTML = `
      <div class="pecas-resumo">
        ${resumoRapido
          .map(
            (r) => `<div class="glass pecas-resumo-card">
            <div class="value">${r.txt}</div>
            <div class="hint">${r.sub}</div>
          </div>`
          )
          .join("")}
      </div>
      <div class="pecas-grupos">
        ${grupos
          .map((g) => {
            const rows = g.itens
              .map((p) => {
                const zero = !p.qtd ? "zero" : "";
                return `<tr class="${zero}">
                  <td>${escapeHtml(p.peca)}</td>
                  <td class="mono">${escapeHtml(p.tamanho)}</td>
                  <td class="mono qtd">${Number(p.qtd).toLocaleString("pt-BR")}</td>
                  <td class="mono">${escapeHtml(p.unidade)}</td>
                  <td class="frase">${escapeHtml(frasePeca(p))}</td>
                </tr>`;
              })
              .join("");
            return `<article class="glass pecas-grupo">
              <header>
                <h3>${escapeHtml(g.nome)}</h3>
                <span class="badge ok">${g.total.toLocaleString("pt-BR")} un</span>
              </header>
              <div class="pecas-table-wrap">
                <table class="pecas-table">
                  <thead>
                    <tr>
                      <th>Peça</th>
                      <th>Tamanho</th>
                      <th>Qtd</th>
                      <th>Un</th>
                      <th>Leitura</th>
                    </tr>
                  </thead>
                  <tbody>${rows}</tbody>
                </table>
              </div>
            </article>`;
          })
          .join("")}
      </div>
    `;
  }

  function frasePeca(p) {
    const q = Number(p.qtd) || 0;
    const nome = pluralPeca(p.peca, q);
    if (!q) return `— sem ${nome} ${p.tamanho}`;
    return `${q.toLocaleString("pt-BR")} ${nome} de ${p.tamanho}`;
  }

  function pluralPeca(nome, qtd) {
    const map = {
      Módulo: ["módulo", "módulos"],
      Base: ["base", "bases"],
      Prateleira: ["prateleira", "prateleiras"],
      Braço: ["braço", "braços"],
      Barra: ["barra", "barras"],
      "Gancho": ["gancho", "ganchos"],
      "Gancho vassoura": ["gancho vassoura", "ganchos vassoura"],
      "Anteparo aramado": ["anteparo aramado", "anteparos aramados"],
      "Anteparo acrílico": ["anteparo acrílico", "anteparos acrílicos"],
      "Cesto carvão": ["cesto carvão", "cestos carvão"],
      "Divisor de cesto": ["divisor de cesto", "divisores de cesto"],
      Cabeceira: ["cabeceira", "cabeceiras"],
      "Porta etiqueta": ["porta etiqueta", "portas etiqueta"],
      Cesto: ["cesto", "cestos"],
    };
    const pair = map[nome] || [nome.toLowerCase(), nome.toLowerCase() + "s"];
    return qtd === 1 ? pair[0] : pair[1];
  }

  function chipHtml(g) {
    const active = state.selectedId === g.id ? "active" : "";
    const partial = g.preenchimento !== "completo" ? "parcial" : "";
    return `<button class="gondola-chip tipo-${g.tipo} ${active} ${partial}"
      data-id="${g.id}" title="${escapeAttr(g.categorias || g.tipoLabel)}">
      <span class="gid">${g.id}</span>
      <span class="gmeta">${g.totalMods} mód.</span>
    </button>`;
  }

  function renderFloor() {
    const list = filteredGondolas();
    const byTipo = (tipo) => list.filter((g) => g.tipo === tipo);

    const zones = [
      { key: "go", title: "Mercearia · GO", items: byTipo("GO"), cols: 3 },
      { key: "gb", title: "Parede · GB", items: byTipo("GB"), cols: 1 },
      { key: "br", title: "Bebidas · BR / BA", items: [...byTipo("BR"), ...byTipo("BA")], cols: 3 },
      {
        key: "bp",
        title: "Padaria · BP / GP / MQ / GS",
        items: [...byTipo("BP"), ...byTipo("GP"), ...byTipo("MQ"), ...byTipo("GS")],
        cols: 4,
      },
      {
        key: "ic",
        title: "Ilhas · IC",
        items: byTipo("IC"),
        cols: 4,
      },
      {
        key: "mh",
        title: "Hortifruti · MH / MM / MS / EX / EC",
        items: [...byTipo("MH"), ...byTipo("MM"), ...byTipo("MS"), ...byTipo("EX"), ...byTipo("EC"), ...byTipo("GC")],
        cols: 4,
      },
    ];

    els.floorMap.innerHTML = zones
      .map((z) => {
        if (!z.items.length) return "";
        return `<div class="floor-zone zone-${z.key}">
          <div class="floor-zone-title">${z.title}</div>
          <div class="floor-zone-grid" style="--cols:${z.cols}">
            ${z.items.map(chipHtml).join("")}
          </div>
        </div>`;
      })
      .join("");

    els.floorMap.querySelectorAll("[data-id]").forEach((btn) => {
      btn.addEventListener("click", () => selectGondola(btn.dataset.id));
    });
  }

  function renderCategories() {
    const list = filteredGondolas()
      .slice()
      .sort((a, b) => typeOrder.indexOf(a.tipo) - typeOrder.indexOf(b.tipo) || a.id.localeCompare(b.id));

    els.catList.innerHTML = list
      .map((g) => {
        const score = fillScore(g);
        const st = statusOf(g);
        const color =
          st === "ok" ? "var(--heatmap-green)" : st === "warn" ? "var(--heatmap-yellow)" : "var(--heatmap-red)";
        const active = state.selectedId === g.id ? "active" : "";
        return `<article class="cat-card glass ${active}" data-id="${g.id}">
          <div class="top">
            <div>
              <h3>${g.id}</h3>
              <div class="sub">${g.tipoLabel} · ${(g.categorias || "—").slice(0, 64)}</div>
            </div>
            <div class="score" style="color:${color}">${score}%</div>
          </div>
          <div class="bar"><i style="width:${score}%;background:${color}"></i></div>
          <div class="foot">
            <span>${g.totalMods} módulos</span>
            <span>${g.totalPrat} prat. · ${g.totalBracos} braços</span>
          </div>
        </article>`;
      })
      .join("");

    els.catList.querySelectorAll("[data-id]").forEach((el) => {
      el.addEventListener("click", () => selectGondola(el.dataset.id));
    });
  }

  function modStatus(m) {
    if (m.sintetico || m.prateleiras === 0) return "warn";
    if (m.prateleiras > 0) return "ok";
    return "empty";
  }

  function faceLabel(face, index) {
    if (face && face.lado) return String(face.lado).toUpperCase();
    const raw = String(face && face.id ? face.id : "");
    const m = raw.match(/\.([ABPC])$/i);
    if (m) return m[1].toUpperCase();
    return index === 0 ? "única" : String(index + 1);
  }

  function bomOf(m) {
    return m && m.bom ? m.bom : {};
  }

  function tdN(v) {
    const n = Number(v) || 0;
    return `<td class="mono">${n ? n : "—"}</td>`;
  }

  function modCellHtml(m) {
    const cls = modStatus(m);
    const shelves = Math.max(0, Math.min(m.prateleiras || 0, 6));
    const shelfBars =
      shelves > 0
        ? `<div class="mod-shelves">${Array.from({ length: shelves }, () => "<i></i>").join("")}</div>`
        : '<div class="mod-shelves empty-shelves"><i></i><i></i><i></i></div>';
    const title = `${m.id} · ${(m.categoria || "").slice(0, 40)} · ${m.prateleiras} prat. · ${m.bracos} braços`;
    return `<div class="mod-cell ${cls}" title="${escapeAttr(title)}">${shelfBars}</div>`;
  }

  /** Uma gôndola física: faces A/B como lados da mesma unidade (não dois equipamentos). */
  function gondolaUnitPreview(g, opts = {}) {
    const faces = (g.faces || []).filter((f) => f && f.modulos && f.modulos.length);
    if (!faces.length) {
      return `<div class="gondola-unit faces-0"><div class="gondola-face"><div class="face-modules"><div class="mod-cell empty"></div></div></div></div>`;
    }
    const compact = opts.compact ? " compact" : "";
    const n = faces.length;
    const faceBlocks = faces
      .map((face, idx) => {
        const tag = faceLabel(face, idx);
        const mods = face.modulos.map(modCellHtml).join("");
        return `<div class="gondola-face" data-face="${escapeAttr(face.id)}">
          <span class="face-tag" title="Face ${tag}">${tag}</span>
          <div class="face-modules">${mods}</div>
          <span class="face-meta">${face.modulos.length} mód.</span>
        </div>`;
      })
      .join(n > 1 ? '<div class="gondola-spine" title="Costas / estrutura compartilhada" aria-hidden="true"><span>costas</span></div>' : "");

    const wall = n === 1 ? '<div class="gondola-wall" title="Parede / costas" aria-hidden="true"></div>' : "";
    const caption =
      n > 1
        ? `<div class="gondola-caption">1 gôndola · ${n} faces · ${g.totalMods} módulos</div>`
        : `<div class="gondola-caption">1 equipamento · 1 face · ${g.totalMods} módulo(s)</div>`;

    return `<div class="gondola-unit faces-${n}${compact}">${caption}<div class="gondola-body">${wall}${faceBlocks}</div></div>`;
  }

  function escapeAttr(s) {
    return String(s).replace(/"/g, "&quot;").replace(/</g, "&lt;");
  }

  function renderGondolaGrid() {
    const list = filteredGondolas();
    els.gondolaGrid.innerHTML = list
      .map((g) => {
        const st = statusOf(g);
        const active = state.selectedId === g.id ? "active" : "";
        const label = st === "ok" ? "BOM OK" : st === "warn" ? "PARCIAL" : "VAZIO";
        return `<article class="gondola-card glass tipo-${g.tipo} ${active}" data-id="${g.id}">
          <div class="head">
            <div>
              <div class="title">${g.id}</div>
              <div class="sub">${g.categorias || g.tipoLabel}</div>
            </div>
            <div class="badge ${st}">${label}</div>
          </div>
          ${gondolaUnitPreview(g, { compact: true })}
          <div class="card-stats">
            <span>${g.faces.length} face(s)</span>
            <span>${g.totalMods} módulos</span>
            <span>${g.totalPrat} prateleiras</span>
            <span>${g.totalBracos} braços</span>
          </div>
        </article>`;
      })
      .join("");

    els.gondolaGrid.querySelectorAll("[data-id]").forEach((el) => {
      el.addEventListener("click", () => selectGondola(el.dataset.id));
    });
  }

  function setActiveNav(view) {
    state.view = view;
    els.navButtons.forEach((b) => b.classList.toggle("active", b.dataset.nav === view));
  }

  function renderDetailEmpty() {
    const list = filteredGondolas();
    const picks = list
      .map(
        (g) => `<button type="button" class="detail-pick" data-id="${g.id}">
          <span class="pid">${g.id}</span>
          <span class="pmeta">${g.totalMods} mód. · ${g.faces.length} face(s)</span>
        </button>`
      )
      .join("");
    els.detail.classList.add("open");
    els.detailBody.innerHTML = `
      <div class="detail-empty glass">
        <h3>Selecione um equipamento</h3>
        <p>O menu Detalhe abre a ficha da gôndola (faces A/B, módulos e BOM). Escolha abaixo ou clique em qualquer card/planta.</p>
        <div class="detail-pick-grid">${picks || "<p>Nenhum equipamento no filtro atual.</p>"}</div>
      </div>
    `;
    els.detailBody.querySelectorAll(".detail-pick").forEach((btn) => {
      btn.addEventListener("click", () => selectGondola(btn.dataset.id));
    });
  }

  function renderDetail(g) {
    if (!g) {
      if (state.view === "detalhe") renderDetailEmpty();
      else {
        els.detail.classList.remove("open");
        els.detailBody.innerHTML = "";
      }
      return;
    }
    els.detail.classList.add("open");
    const facesHtml = g.faces
      .map((face, idx) => {
        const tag = faceLabel(face, idx);
        const rows = face.modulos
          .map((m) => {
            const b = bomOf(m);
            const dimTxt = [m.altura, m.largura || m.dim, m.profundidade].filter(Boolean).join(" × ") || m.dim || "—";
            return `<tr>
              <td class="mono">${escapeHtml(m.id)}</td>
              <td>${escapeHtml((m.categoria || "—").slice(0, 80))}</td>
              <td class="mono">${escapeHtml(dimTxt)}</td>
              ${tdN(m.prateleiras)}
              ${tdN(b.bracos040)}
              ${tdN(b.bracos035)}
              ${tdN(b.bracos060)}
              ${tdN(b.barra100)}
              ${tdN(b.barras130)}
              ${tdN(b.ganchos3530)}
              ${tdN(b.ganchos3025)}
              ${tdN(b.ganchos4035)}
              ${tdN(b.ganchosVassoura)}
              ${tdN(b.anteparoAramado)}
              ${tdN(b.anteparoAcrilico)}
              ${tdN(b.cestos100)}
              ${tdN(b.cestos130)}
              ${tdN(b.cestoVinhos)}
              ${tdN(b.cestoCarvao)}
              ${tdN(b.divisorCesto100)}
              ${tdN(b.divisorCesto130)}
              ${tdN(b.cabeceira)}
              <td>${escapeHtml((m.observacao || m.especificacao || "—").slice(0, 90))}</td>
            </tr>`;
          })
          .join("");
        return `<section class="face-panel glass">
          <h3>Face ${tag} <span class="face-id-muted">${face.id}</span></h3>
          <div class="face-cat">${escapeHtml(face.categoria || "—")} · linear: ${escapeHtml(face.linear || "—")} · ${face.headerMods} módulos${face.tpMobiliario ? " · " + escapeHtml(face.tpMobiliario) : ""}</div>
          <div class="viz-face-strip">
            <div class="caption">Módulos desta face (lado ${tag} da mesma gôndola)</div>
            <div class="face-modules strip">${face.modulos.map(modCellHtml).join("")}</div>
          </div>
          <div style="overflow:auto;margin-top:14px">
            <table class="module-table">
              <thead>
                <tr>
                  <th>Módulo</th><th>Categoria</th><th>Dimensão</th>
                  <th>Prat.</th><th>Br 40</th><th>Br 35</th><th>Br 60</th>
                  <th>Barra 1,00</th><th>Barra 1,30</th>
                  <th>G 35×30</th><th>G 30×25</th><th>G 40×35</th><th>Vass.</th>
                  <th>Ant. aram.</th><th>Ant. acr.</th>
                  <th>Cesto 1,00</th><th>Cesto 1,30</th><th>Vinhos</th><th>Carvão</th>
                  <th>Div. 1,00</th><th>Div. 1,30</th><th>Cab.</th><th>Observações</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        </section>`;
      })
      .join("");

    els.detailBody.innerHTML = `
      <div class="detail-header">
        <div>
          <h2>${g.id} · ${g.tipoLabel}</h2>
          <p>${escapeHtml(g.setor || "")} — ${escapeHtml(g.categorias || "")}</p>
          <div class="card-stats" style="margin-top:10px">
            <span class="pill ${g.preenchimento}">${g.preenchimento}</span>
            <span>${g.totalMods} módulos</span>
            <span>${g.totalPrat} prateleiras</span>
            <span>${g.totalBracos} braços</span>
            <span>${g.totalGanchos || 0} ganchos</span>
            <span>${g.totalCestos || 0} cestos</span>
            <span>${g.faces.length} face(s)</span>
          </div>
        </div>
        <div class="detail-actions">
          <button class="btn" id="btnCloseDetail"><span class="material-symbols-outlined" style="font-size:18px">close</span> Fechar</button>
        </div>
      </div>
      <div class="detail-unit glass">
        <div class="caption">Vista da unidade — faces como lados da mesma gôndola</div>
        ${gondolaUnitPreview(g)}
      </div>
      <div class="faces">${facesHtml}</div>
    `;
    document.getElementById("btnCloseDetail").addEventListener("click", () => {
      state.selectedId = null;
      if (state.view === "detalhe") {
        renderAll();
        els.detail.scrollIntoView({ behavior: "smooth", block: "start" });
      } else {
        renderAll();
      }
    });
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function selectGondola(id) {
    state.selectedId = id;
    setActiveNav("detalhe");
    renderAll();
    els.detail.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function openDetailView() {
    setActiveNav("detalhe");
    const list = filteredGondolas();
    const stillValid = list.some((g) => g.id === state.selectedId);
    if (!stillValid) state.selectedId = list[0]?.id || null;
    renderAll();
    els.detail.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function renderAll() {
    renderKpis();
    renderPecas();
    renderFloor();
    renderCategories();
    renderGondolaGrid();
    const g = DATA.gondolas.find((x) => x.id === state.selectedId);
    renderDetail(g || null);
  }

  function openHelp() {
    if (!els.helpModal) return;
    els.helpModal.hidden = false;
    document.body.classList.add("help-open");
    (els.btnCloseHelp || els.helpModal).focus();
  }

  function closeHelp() {
    if (!els.helpModal) return;
    els.helpModal.hidden = true;
    document.body.classList.remove("help-open");
    if (els.btnHelp) els.btnHelp.focus();
  }

  // Events
  els.filterTipo.addEventListener("change", (e) => {
    state.filterTipo = e.target.value;
    renderAll();
  });
  els.search.addEventListener("input", (e) => {
    state.query = e.target.value;
    renderAll();
  });
  els.navButtons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const view = btn.dataset.nav;
      if (view === "detalhe") {
        openDetailView();
        return;
      }
      setActiveNav(view);
      const map = {
        planta: "section-planta",
        gondolas: "section-gondolas",
        resumo: "section-kpis",
        pecas: "section-pecas",
      };
      const target = document.getElementById(map[view] || "section-planta");
      if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });

  if (els.btnHelp) els.btnHelp.addEventListener("click", openHelp);
  if (els.btnCloseHelp) els.btnCloseHelp.addEventListener("click", closeHelp);
  if (els.helpModal) {
    els.helpModal.addEventListener("click", (e) => {
      if (e.target === els.helpModal) closeHelp();
    });
  }
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && els.helpModal && !els.helpModal.hidden) {
      e.preventDefault();
      closeHelp();
    }
  });

  renderAll();
})();
