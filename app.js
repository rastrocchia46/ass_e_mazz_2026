(() => {
  "use strict";

  const state = {
    data: null,
    activeView: "dashboard",
    rankingRound: 1,
    tableRound: 1,
    refreshInFlight: false,
  };

  const POINTS_BY_POSITION = { 1: 9, 2: 6, 3: 4, 4: 2, 5: 1 };
  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];


  const normalise = (value) => String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();

  const toNumber = (value) => {
    if (value === "" || value === null || value === undefined) return null;
    const number = Number(String(value).replace(",", "."));
    return Number.isFinite(number) ? number : null;
  };

  const escapeHtml = (value) => String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");

  function sheetRows(workbook, name) {
    const sheet = workbook.Sheets[name];
    if (!sheet) throw new Error(`Foglio Excel mancante: ${name}`);
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
  }

  function locateHeader(rows, expected) {
    const expectedNormalised = expected.map(normalise);
    const index = rows.findIndex((row) => {
      const values = row.map(normalise);
      return expectedNormalised.every((header) => values.includes(header));
    });
    if (index < 0) throw new Error(`Intestazione non trovata: ${expected.join(", ")}`);
    return index;
  }

  function rowsToObjects(rows, headerIndex) {
    const headers = rows[headerIndex].map(normalise);
    return rows.slice(headerIndex + 1)
      .filter((row) => row.some((cell) => cell !== ""))
      .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));
  }

  function rowsToTournamentData(configRows, playersRows, resultRows, source = "excel") {
    const config = {};
    for (const row of configRows) {
      const key = normalise(row[0]);
      if (key && row.length > 1) config[key] = row[1];
    }

    const playersCount = toNumber(config["numero giocatori"]) || 65;
    const playersHeader = locateHeader(playersRows, ["ID", "Nome e cognome"]);
    const playerObjects = rowsToObjects(playersRows, playersHeader);
    const players = playerObjects
      .map((row) => {
        const id = toNumber(row.id);
        if (!id || id > playersCount) return null;
        return {
          id,
          name: String(row["nome e cognome"] || "").trim(),
          nickname: String(row["soprannome (facoltativo)"] || "").trim(),
        };
      })
      .filter(Boolean);

    const resultsHeader = locateHeader(resultRows, ["Turno", "Tavolo", "Giocatore ID", "Posizione"]);
    const resultObjects = rowsToObjects(resultRows, resultsHeader);
    const results = resultObjects
      .map((row) => ({
        turn: toNumber(row.turno),
        table: String(row.tavolo || "").trim().toUpperCase(),
        seat: toNumber(row.sedia),
        playerId: toNumber(row["giocatore id"]),
        gamePoints: toNumber(row["punti partita"]),
        position: toNumber(row.posizione),
        sole: toNumber(row.sole) ?? 0,
        note: String(row.note || "").trim(),
      }))
      .filter((row) => row.turn && row.table && row.playerId);

    return {
      source,
      config: {
        title: String(config["nome torneo"] || "Torneo Ass 'e Mazz 2026"),
        playersCount,
        cutoff: toNumber(config["posti qualificazione"]) || 25,
        phase: String(config.fase || "Eliminatorie"),
      },
      players,
      results,
    };
  }

  function loadFromAppsScript() {
    return new Promise((resolve, reject) => {
      const configuredUrl = String(window.TORNEO_CONFIG?.appsScriptUrl || "").trim();
      if (!configuredUrl || configuredUrl.includes("INCOLLA_QUI")) {
        reject(new Error("URL Apps Script non configurato."));
        return;
      }

      const callbackName = `__torneoData_${Date.now()}_${Math.random().toString(36).slice(2)}`;
      const separator = configuredUrl.includes("?") ? "&" : "?";
      const script = document.createElement("script");
      let settled = false;

      const cleanup = () => {
        script.remove();
        try { delete window[callbackName]; } catch (_) { window[callbackName] = undefined; }
      };

      const timeout = window.setTimeout(() => {
        if (settled) return;
        settled = true;
        cleanup();
        reject(new Error("Tempo scaduto durante la lettura di Google Sheets."));
      }, 15000);

      window[callbackName] = (payload) => {
        if (settled) return;
        settled = true;
        window.clearTimeout(timeout);
        cleanup();
        if (!payload || !payload.config || !Array.isArray(payload.players) || !Array.isArray(payload.results)) {
          reject(new Error("Risposta Apps Script non valida."));
          return;
        }
        payload.source = "apps-script";
        resolve(payload);
      };

      script.onerror = () => {
        if (settled) return;
        settled = true;
        window.clearTimeout(timeout);
        cleanup();
        reject(new Error("Impossibile contattare la Web app Apps Script."));
      };

      script.src = `${configuredUrl}${separator}callback=${encodeURIComponent(callbackName)}&_=${Date.now()}`;
      script.async = true;
      document.head.appendChild(script);
    });
  }

  function workbookToData(workbook) {
    const configRows = sheetRows(workbook, "Config");
    const playersRows = sheetRows(workbook, "Giocatori");

    const config = {};
    for (const row of configRows) {
      const key = normalise(row[0]);
      if (key && row.length > 1) config[key] = row[1];
    }
    const playersCount = toNumber(config["numero giocatori"]) || 65;
    const resultRows = sheetRows(workbook, `Risultati_${playersCount}`);
    return rowsToTournamentData(configRows, playersRows, resultRows, "excel");
  }

  async function loadTournamentData() {
    let appsScriptError = null;
    let excelError = null;

    try {
      return await loadFromAppsScript();
    } catch (error) {
      appsScriptError = error;
    }

    try {
      if (!window.XLSX) throw new Error("La libreria necessaria per leggere il file Excel non è disponibile.");
      const response = await fetch(`data/torneo.xlsx?v=${Date.now()}`, { cache: "no-store" });
      if (!response.ok) throw new Error(`File Excel non disponibile (${response.status}).`);
      const buffer = await response.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const data = workbookToData(workbook);
      data.appsScriptError = appsScriptError?.message || "Errore sconosciuto";
      return data;
    } catch (error) {
      excelError = error;
    }

    try {
      const response = await fetch(`data/torneo.json?v=${Date.now()}`, { cache: "no-store" });
      if (!response.ok) throw new Error("Dati di riserva non disponibili.");
      const fallback = await response.json();
      fallback.source = "fallback";
      fallback.appsScriptError = appsScriptError?.message || "Errore sconosciuto";
      fallback.excelError = excelError?.message || "Errore sconosciuto";
      return fallback;
    } catch (fallbackError) {
      throw new Error(`${appsScriptError?.message || "Errore Apps Script"} ${excelError?.message || "Errore Excel"} ${fallbackError.message}`);
    }
  }

  function buildPlayerMap(data) {
    const byId = new Map(data.players.map((player) => [Number(player.id), player]));
    const map = new Map();

    for (let id = 1; id <= data.config.playersCount; id += 1) {
      const stored = byId.get(id) || {};
      const label = stored.name
        ? stored.nickname ? `${stored.name} “${stored.nickname}”` : stored.name
        : stored.nickname || `Giocatore ${id}`;
      map.set(id, { id, name: label, rawName: stored.name || "", nickname: stored.nickname || "" });
    }

    return map;
  }

  function rowClassPoints(row) {
    if (!POINTS_BY_POSITION[row.position]) return 0;
    return POINTS_BY_POSITION[row.position] + (Math.max(0, row.sole || 0) * 2);
  }

  function isCompletedRow(row) {
    return Number.isInteger(row.position) && row.position >= 1 && row.position <= 5;
  }

  function isStartedRow(row) {
    return isCompletedRow(row)
      || row.gamePoints !== null
      || (row.note && row.note.length > 0)
      || (row.sole && row.sole > 0);
  }

  function tableGroups(turn) {
    const groups = new Map();
    for (const row of state.data.results.filter((item) => item.turn === turn)) {
      if (!groups.has(row.table)) groups.set(row.table, []);
      groups.get(row.table).push(row);
    }

    return [...groups.entries()]
      .sort(([a], [b]) => a.localeCompare(b, "it"))
      .map(([table, rows]) => {
        const sortedRows = [...rows].sort((a, b) => a.seat - b.seat);
        const validPositions = sortedRows
          .filter(isCompletedRow)
          .map((row) => row.position)
          .sort((a, b) => a - b);
        const completed = sortedRows.length === 5 && validPositions.join(",") === "1,2,3,4,5";
        return {
          table,
          rows: sortedRows,
          completed,
          started: sortedRows.some(isStartedRow),
        };
      });
  }

  function globalTableStats() {
    const all = [1, 2, 3].flatMap((turn) => tableGroups(turn));
    return {
      total: all.length,
      completed: all.filter((group) => group.completed).length,
      live: all.filter((group) => group.started && !group.completed).length,
      todo: all.filter((group) => !group.started).length,
    };
  }

  function currentResultsRound() {
    const startedTurns = [1, 2, 3].filter((turn) => tableGroups(turn).some((group) => group.started));
    return startedTurns.at(-1) || 0;
  }

  function currentStage() {
    const round = currentResultsRound();
    if (!round) return { short: "—", label: "In attesa del primo turno" };

    const groups = tableGroups(round);
    const allCompleted = groups.length > 0 && groups.every((group) => group.completed);
    if (round === 3 && allCompleted) return { short: "Fine", label: "Fase eliminatoria conclusa" };
    if (allCompleted) return { short: `T${round}`, label: `Turno ${round} completato` };
    return { short: `T${round}`, label: `Turno ${round} in corso` };
  }

  function completedTableKeySet() {
    return new Set(
      [1, 2, 3]
        .flatMap((turn) => tableGroups(turn).map((group) => ({ turn, ...group })))
        .filter((group) => group.completed)
        .map((group) => `${group.turn}-${group.table}`),
    );
  }

  function compareStandings(a, b) {
    return (b.points - a.points)
      || (b.sole - a.sole)
      || (a.bestPosition - b.bestPosition)
      || (a.minGamePoints - b.minGamePoints)
      || (a.totalGamePoints - b.totalGamePoints)
      || (a.id - b.id);
  }

  function sameOfficialCriteria(a, b) {
    return a.points === b.points
      && a.sole === b.sole
      && a.bestPosition === b.bestPosition
      && a.minGamePoints === b.minGamePoints
      && a.totalGamePoints === b.totalGamePoints;
  }

  function computeStandings(maxTurn) {
    const data = state.data;
    const playerMap = buildPlayerMap(data);
    const completedTables = completedTableKeySet();
    const standings = [...playerMap.values()].map((player) => ({
      ...player,
      played: 0,
      points: 0,
      sole: 0,
      bestPosition: Infinity,
      minGamePoints: Infinity,
      totalGamePoints: 0,
      scoreDataComplete: true,
      roundPoints: { 1: 0, 2: 0, 3: 0 },
    }));
    const byId = new Map(standings.map((entry) => [entry.id, entry]));

    for (const row of data.results) {
      if (row.turn > maxTurn || !isCompletedRow(row) || !completedTables.has(`${row.turn}-${row.table}`)) continue;
      const entry = byId.get(row.playerId);
      if (!entry) continue;

      const classPoints = rowClassPoints(row);
      entry.played += 1;
      entry.points += classPoints;
      entry.roundPoints[row.turn] += classPoints;
      entry.sole += Math.max(0, row.sole || 0);
      entry.bestPosition = Math.min(entry.bestPosition, row.position);

      if (row.gamePoints === null) {
        entry.scoreDataComplete = false;
      } else {
        entry.minGamePoints = Math.min(entry.minGamePoints, row.gamePoints);
        entry.totalGamePoints += row.gamePoints;
      }
    }

    for (const entry of standings) {
      if (!entry.played) {
        entry.bestPosition = Infinity;
        entry.minGamePoints = Infinity;
        entry.totalGamePoints = Infinity;
      } else if (!entry.scoreDataComplete) {
        entry.minGamePoints = Infinity;
        entry.totalGamePoints = Infinity;
      }
    }

    standings.sort(compareStandings);

    let groupStart = 0;
    while (groupStart < standings.length) {
      let groupEnd = groupStart;
      while (groupEnd + 1 < standings.length && sameOfficialCriteria(standings[groupStart], standings[groupEnd + 1])) {
        groupEnd += 1;
      }

      const groupHasResults = standings[groupStart].played > 0;
      const perfectTie = groupHasResults && groupEnd > groupStart;
      const cutoffTie = perfectTie && groupStart < data.config.cutoff && groupEnd >= data.config.cutoff;
      const displayRank = groupStart + 1;

      for (let index = groupStart; index <= groupEnd; index += 1) {
        standings[index].displayRank = displayRank;
        standings[index].perfectTie = perfectTie;
        standings[index].cutoffTie = cutoffTie;
        standings[index].qualified = groupHasResults && !cutoffTie && groupEnd < data.config.cutoff;
      }
      groupStart = groupEnd + 1;
    }

    return standings;
  }

  function displayMetric(value) {
    return Number.isFinite(value) ? value : "—";
  }

  function renderHeader() {
    const { config, source, appsScriptError, excelError } = state.data;
    $("#tournament-title").textContent = config.title;
    $("#phase-label").textContent = config.phase;

    const warning = $("#warning-banner");
    warning.classList.add("hidden");

    if (source === "excel") {
      warning.textContent = `Google Sheets non è raggiungibile (${appsScriptError}). Sono visualizzati i dati del file Excel incluso nel sito.`;
      warning.classList.remove("hidden");
    } else if (source === "fallback") {
      warning.textContent = `Google Sheets e il file Excel non sono raggiungibili. È visualizzata la copia dati di riserva inclusa nel sito.`;
      warning.classList.remove("hidden");
    }
  }

  function renderDashboard() {
    const { config } = state.data;
    const stats = globalTableStats();
    const stage = currentStage();
    const currentRound = currentResultsRound();
    const percent = stats.total ? Math.round((stats.completed / stats.total) * 100) : 0;

    $("#stat-players").textContent = config.playersCount;
    $("#stat-cutoff").textContent = `Top ${config.cutoff}`;
    $("#round-status").textContent = stage.label;

    $("#progress-ring").style.setProperty("--progress", percent);
    $("#progress-percent").textContent = `${percent}%`;
    $("#progress-todo").textContent = stats.todo;
    $("#progress-live").textContent = stats.live;
    $("#progress-done").textContent = stats.completed;

    const top = currentRound
      ? computeStandings(currentRound).filter((entry) => entry.played > 0).slice(0, 10)
      : [];
    const container = $("#top-ranking");

    if (!top.length) {
      container.className = "compact-ranking empty-state";
      container.textContent = "Nessun tavolo ancora completato.";
      return;
    }

    container.className = "compact-ranking";
    container.innerHTML = top.map((entry) => `
      <div class="rank-line">
        <span class="rank-number">${entry.displayRank}</span>
        <div class="rank-player">
          <strong>${escapeHtml(entry.name)}</strong>
          <small>${entry.played} ${entry.played === 1 ? "partita" : "partite"} · ${entry.sole} ${entry.sole === 1 ? "sóla" : "sóle"}</small>
        </div>
        <span class="rank-points">${entry.points} pt</span>
      </div>
    `).join("");
  }

  function renderRanking() {
    const round = state.rankingRound;
    const search = normalise($("#ranking-search").value);
    const standings = computeStandings(round);
    const filtered = standings.filter((entry) => !search
      || normalise(entry.name).includes(search)
      || String(entry.id).includes(search));

    const tieAtCutoff = standings.some((entry) => entry.cutoffTie);
    const notice = $("#tie-notice");
    if (tieAtCutoff) {
      notice.textContent = `Una perfetta parità coinvolge la posizione di qualificazione n. ${state.data.config.cutoff}: è necessario il sorteggio previsto dal regolamento.`;
      notice.classList.remove("hidden");
    } else {
      notice.classList.add("hidden");
    }

    const body = $("#ranking-body");
    body.innerHTML = filtered.map((entry) => {
      const rowClass = entry.cutoffTie || entry.perfectTie ? "playoff" : entry.qualified ? "qualified" : "";
      const cutoffClass = entry.displayRank === state.data.config.cutoff ? " cutoff" : "";
      const tieBadge = entry.perfectTie ? '<span class="badge draw">sorteggio</span>' : "";
      return `
        <tr class="${rowClass}${cutoffClass}">
          <td><span class="position-medal ${entry.displayRank <= 3 ? "top" : ""}">${entry.displayRank}</span></td>
          <td class="player-cell"><strong>${escapeHtml(entry.name)}${tieBadge}</strong><small>ID ${entry.id}</small></td>
          <td class="center">${entry.played}</td>
          <td class="center round-col">${entry.roundPoints[1] || "—"}</td>
          <td class="center round-col">${round >= 2 ? entry.roundPoints[2] || "—" : "—"}</td>
          <td class="center round-col">${round >= 3 ? entry.roundPoints[3] || "—" : "—"}</td>
          <td class="center">${entry.sole}</td>
          <td class="center score-main">${entry.points}</td>
          <td class="center tie-col">${displayMetric(entry.bestPosition)}</td>
          <td class="center tie-col">${displayMetric(entry.minGamePoints)}</td>
          <td class="center tie-col">${displayMetric(entry.totalGamePoints)}</td>
        </tr>`;
    }).join("");

    if (!filtered.length) {
      body.innerHTML = '<tr><td colspan="11" class="center">Nessun giocatore trovato.</td></tr>';
    }
  }

  function renderTables() {
    const turn = state.tableRound;
    const playerMap = buildPlayerMap(state.data);
    const groups = tableGroups(turn);
    const complete = groups.filter((group) => group.completed).length;
    const live = groups.filter((group) => group.started && !group.completed).length;
    const todo = groups.length - complete - live;
    $("#tables-summary").textContent = `${groups.length} tavoli · ${complete} completati · ${live} in corso · ${todo} da giocare`;

    $("#tables-grid").innerHTML = groups.map((group) => {
      const orderedRows = group.completed
        ? [...group.rows].sort((a, b) => a.position - b.position)
        : group.rows;
      const statusClass = group.completed ? "done" : group.started ? "live" : "";
      const statusLabel = group.completed ? "Concluso" : group.started ? "In corso" : "Da giocare";

      return `
        <article class="table-card">
          <header class="table-card-header">
            <h3>Tavolo ${escapeHtml(group.table)}</h3>
            <span class="table-status ${statusClass}">${statusLabel}</span>
          </header>
          <div class="table-players">
            ${orderedRows.map((row) => {
              const player = playerMap.get(row.playerId);
              const completed = isCompletedRow(row);
              const badge = completed ? `${row.position}°` : row.seat;
              const details = completed
                ? `${row.gamePoints ?? "—"} punti${row.sole ? ` · ${row.sole} ${row.sole === 1 ? "sóla" : "sóle"}` : ""}`
                : `Sedia ${row.seat}`;

              return `
                <div class="table-player ${row.position === 1 ? "winner" : ""}">
                  <span class="seat-or-place">${badge}</span>
                  <div class="table-player-name">
                    <strong>${escapeHtml(player?.name || `Giocatore ${row.playerId}`)}</strong>
                    <small>${details}</small>
                  </div>
                  <div class="table-player-result">
                    <strong>${completed ? `${rowClassPoints(row)} pt` : "—"}</strong>
                    <small>classifica</small>
                  </div>
                </div>`;
            }).join("")}
          </div>
        </article>`;
    }).join("");
  }

  function setView(view) {
    state.activeView = view;
    $$(".view").forEach((section) => section.classList.toggle("active", section.id === view));
    $$(".nav-button").forEach((button) => {
      const active = button.dataset.view === view;
      button.classList.toggle("active", active);
      if (active) button.setAttribute("aria-current", "page");
      else button.removeAttribute("aria-current");
    });
    history.replaceState(null, "", `#${view}`);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function bindEvents() {
    $$(".nav-button").forEach((button) => button.addEventListener("click", () => setView(button.dataset.view)));
    $$("[data-open-view]").forEach((button) => button.addEventListener("click", () => setView(button.dataset.openView)));

    $("#ranking-round").addEventListener("change", (event) => {
      state.rankingRound = Number(event.target.value);
      renderRanking();
    });
    $("#ranking-search").addEventListener("input", renderRanking);

    $$(".round-tab").forEach((button) => button.addEventListener("click", () => {
      state.tableRound = Number(button.dataset.round);
      $$(".round-tab").forEach((item) => item.classList.toggle("active", item === button));
      renderTables();
    }));
  }

  function renderAll() {
    renderHeader();
    renderDashboard();
    renderRanking();
    renderTables();
  }


  async function refreshFromSource() {
    if (state.refreshInFlight) return;
    state.refreshInFlight = true;

    try {
      const freshData = await loadTournamentData();
      state.data = freshData;
      renderAll();
    } catch (error) {
      console.warn("Aggiornamento automatico non riuscito:", error);
    } finally {
      state.refreshInFlight = false;
    }
  }

  function startAutoRefresh() {
    const configuredInterval = Number(window.TORNEO_CONFIG?.refreshMs) || 10000;
    const interval = Math.max(5000, configuredInterval);

    window.setInterval(() => {
      if (!document.hidden) refreshFromSource();
    }, interval);

    document.addEventListener("visibilitychange", () => {
      if (!document.hidden) refreshFromSource();
    });
  }

  async function init() {
    bindEvents();

    try {
      state.data = await loadTournamentData();
      const detectedRound = currentResultsRound();
      state.rankingRound = detectedRound || 1;
      state.tableRound = detectedRound || 1;
      $("#ranking-round").value = String(state.rankingRound);
      $$(".round-tab").forEach((button) => button.classList.toggle("active", Number(button.dataset.round) === state.tableRound));
      renderAll();
      startAutoRefresh();

      const hashView = location.hash.replace("#", "");
      if (["dashboard", "classifica", "tavoli", "regolamento"].includes(hashView)) setView(hashView);
    } catch (error) {
      const banner = $("#error-banner");
      banner.textContent = `Impossibile caricare i dati del torneo: ${error.message}`;
      banner.classList.remove("hidden");
      console.error(error);
    } finally {
      $("#loading").classList.add("done");
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
