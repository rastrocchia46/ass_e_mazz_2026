(() => {
  "use strict";

  const state = {
    data: null,
    activeView: "dashboard",
    rankingRound: 1,
    tableRound: 1,
    finalStage: "ottavi",
    refreshInFlight: false,
  };

  const POINTS_BY_POSITION = { 1: 9, 2: 6, 3: 4, 4: 2, 5: 1 };
  const FINAL_STAGE_META = {
    ottavi: { label: "Ottavi di finale", short: "Ottavi", tableCount: 5, playerCount: 25, topEach: 3, wildcardPlace: null, wildcardCount: 0 },
    quarti: { label: "Quarti di finale", short: "Quarti", tableCount: 3, playerCount: 15, topEach: 3, wildcardPlace: 4, wildcardCount: 1 },
    semifinali: { label: "Semifinali", short: "Semifinali", tableCount: 2, playerCount: 10, topEach: 2, wildcardPlace: 3, wildcardCount: 1 },
    finale: { label: "Tavolo finale", short: "Finale", tableCount: 1, playerCount: 5, topEach: 1, wildcardPlace: null, wildcardCount: 0 },
  };
  const FINAL_STAGE_ORDER = ["ottavi", "quarti", "semifinali", "finale"];
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

  async function loadTournamentData(options = {}) {
    const allowFallback = options.allowFallback !== false;
    let appsScriptError = null;
    let excelError = null;

    try {
      return await loadFromAppsScript();
    } catch (error) {
      appsScriptError = error;
    }

    if (!allowFallback) {
      throw appsScriptError || new Error("Google Sheets non è raggiungibile.");
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
    if (!eliminationsAreComplete()) {
      if (!round) return { short: "—", label: "In attesa del primo turno" };
      const groups = tableGroups(round);
      const allCompleted = groups.length > 0 && groups.every((group) => group.completed);
      if (allCompleted) return { short: `T${round}`, label: `Turno ${round} completato` };
      return { short: `T${round}`, label: `Turno ${round} in corso` };
    }

    const finals = getFinalsData();
    if (finals.currentStage === "concluso") return { short: "Fine", label: "Torneo concluso" };
    const stage = FINAL_STAGE_META[finals.currentStage];
    if (!stage) return { short: "Finali", label: "Fase eliminatoria conclusa" };
    return { short: stage.short, label: `${stage.label} ${finals[finals.currentStage].started ? "in corso" : "da disputare"}` };
  }

  function eliminationsAreComplete() {
    return [1, 2, 3].every((turn) => {
      const groups = tableGroups(turn);
      return groups.length > 0 && groups.every((group) => group.completed);
    });
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

    const finalComplete = maxTurn === 3 && eliminationsAreComplete();
    let groupStart = 0;
    while (groupStart < standings.length) {
      let groupEnd = groupStart;
      while (groupEnd + 1 < standings.length && sameOfficialCriteria(standings[groupStart], standings[groupEnd + 1])) {
        groupEnd += 1;
      }

      const groupHasResults = standings[groupStart].played > 0;
      const perfectTie = finalComplete && groupHasResults && groupEnd > groupStart;
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

  function normaliseFinalRow(row, seedByPlayer = new Map()) {
    const playerId = toNumber(row?.playerId);
    return {
      table: String(row?.table || "").trim().toUpperCase(),
      seat: toNumber(row?.seat),
      seed: toNumber(row?.seed) || seedByPlayer.get(playerId) || null,
      playerId,
      gamePoints: toNumber(row?.gamePoints),
      position: toNumber(row?.position),
      note: String(row?.note || "").trim(),
    };
  }

  function groupFinalRows(rows) {
    const groups = new Map();
    for (const row of rows || []) {
      if (!row.table || !row.playerId) continue;
      if (!groups.has(row.table)) groups.set(row.table, []);
      groups.get(row.table).push(row);
    }
    for (const groupRows of groups.values()) {
      groupRows.sort((a, b) => (a.seat || 99) - (b.seat || 99));
    }
    return groups;
  }

  function finalTableIsComplete(rows) {
    if (!rows || rows.length !== 5) return false;
    const positions = rows.filter(isCompletedRow).map((row) => row.position).sort((a, b) => a - b);
    return positions.join(",") === "1,2,3,4,5";
  }

  function finalStageIsComplete(rows, tableCount) {
    const groups = groupFinalRows(rows);
    return groups.size === tableCount && [...groups.values()].every(finalTableIsComplete);
  }

  function finalStageIsStarted(rows) {
    return (rows || []).some((row) => isCompletedRow(row)
      || row.gamePoints !== null
      || (row.note && row.note.length > 0));
  }

  function compareFinalPerformance(a, b) {
    const aPoints = a.gamePoints === null ? Number.POSITIVE_INFINITY : a.gamePoints;
    const bPoints = b.gamePoints === null ? Number.POSITIVE_INFINITY : b.gamePoints;
    return (a.position - b.position)
      || (aPoints - bPoints)
      || ((a.seed || 999) - (b.seed || 999))
      || (a.playerId - b.playerId);
  }

  function getFinalQualifiers(stageKey, rows) {
    const meta = FINAL_STAGE_META[stageKey];
    if (!meta || !finalStageIsComplete(rows, meta.tableCount)) return [];

    const qualifiers = [];
    const wildcards = [];
    const groups = groupFinalRows(rows);
    for (const [table, groupRows] of [...groups.entries()].sort(([a], [b]) => a.localeCompare(b, "it"))) {
      for (const row of groupRows) {
        const candidate = { ...row, sourceTable: table };
        if (row.position <= meta.topEach) qualifiers.push(candidate);
        if (meta.wildcardPlace && row.position === meta.wildcardPlace) wildcards.push(candidate);
      }
    }
    wildcards.sort(compareFinalPerformance);
    return qualifiers.concat(wildcards.slice(0, meta.wildcardCount));
  }

  function buildOttaviRows(standings) {
    const bySeed = new Map(standings.slice(0, 25).map((entry, index) => [index + 1, { ...entry, seed: index + 1 }]));
    const seedTables = [
      { table: "A", seeds: [1, 10, 11, 20, 21] },
      { table: "B", seeds: [2, 9, 12, 19, 22] },
      { table: "C", seeds: [3, 8, 13, 18, 23] },
      { table: "D", seeds: [4, 7, 14, 17, 24] },
      { table: "E", seeds: [5, 6, 15, 16, 25] },
    ];

    return seedTables.flatMap((group) => group.seeds.map((seed, seatIndex) => {
      const player = bySeed.get(seed);
      return {
        table: group.table,
        seat: seatIndex + 1,
        seed,
        playerId: player?.id || null,
        gamePoints: null,
        position: null,
        note: "",
      };
    })).filter((row) => row.playerId);
  }

  function finalRowAt(rows, table, position) {
    const tableKey = String(table).trim().toUpperCase();
    const row = rows.find((item) => String(item.table).trim().toUpperCase() === tableKey && item.position === position);
    return row ? { ...row, sourceTable: tableKey } : null;
  }

  function bestFinalRowAtPosition(rows, position) {
    return rows.filter((row) => row.position === position).sort(compareFinalPerformance)[0] || null;
  }

  function makePreviewFinalRow(table, seat, player) {
    return {
      table,
      seat,
      seed: player?.seed || null,
      playerId: player?.playerId || null,
      gamePoints: null,
      position: null,
      note: "",
    };
  }

  function buildQuartiRows(ottaviRows) {
    const structure = [
      { table: "Q1", refs: [["A", 1], ["B", 2], ["C", 3], ["D", 1], ["E", 2]] },
      { table: "Q2", refs: [["A", 2], ["B", 3], ["C", 1], ["D", 2], ["E", 3]] },
      { table: "Q3", refs: [["A", 3], ["B", 1], ["C", 2], ["D", 3], ["E", 1]] },
    ];
    return structure.flatMap((group) => group.refs.map(([table, position], index) =>
      makePreviewFinalRow(group.table, index + 1, finalRowAt(ottaviRows, table, position))))
      .filter((row) => row.playerId);
  }

  function buildSemifinaliRows(quartiRows) {
    const slots = new Map([
      [1, finalRowAt(quartiRows, "Q1", 1)],
      [2, finalRowAt(quartiRows, "Q2", 1)],
      [3, finalRowAt(quartiRows, "Q3", 1)],
      [4, finalRowAt(quartiRows, "Q1", 2)],
      [5, finalRowAt(quartiRows, "Q2", 2)],
      [6, finalRowAt(quartiRows, "Q3", 2)],
      [7, finalRowAt(quartiRows, "Q1", 3)],
      [8, finalRowAt(quartiRows, "Q2", 3)],
      [9, finalRowAt(quartiRows, "Q3", 3)],
      [10, bestFinalRowAtPosition(quartiRows, 4)],
    ]);
    const structure = [
      { table: "S1", slots: [1, 4, 5, 8, 9] },
      { table: "S2", slots: [2, 3, 6, 7, 10] },
    ];
    return structure.flatMap((group) => group.slots.map((slot, index) =>
      makePreviewFinalRow(group.table, index + 1, slots.get(slot))))
      .filter((row) => row.playerId);
  }

  function buildFinaleRows(semifinaliRows) {
    const finalists = [
      finalRowAt(semifinaliRows, "S1", 1),
      finalRowAt(semifinaliRows, "S1", 2),
      finalRowAt(semifinaliRows, "S2", 1),
      finalRowAt(semifinaliRows, "S2", 2),
      bestFinalRowAtPosition(semifinaliRows, 3),
    ];
    return finalists.map((player, index) => makePreviewFinalRow("FINALISSIMA", index + 1, player))
      .filter((row) => row.playerId);
  }

  function makeFinalStage(stageKey, rows, source = "waiting") {
    const meta = FINAL_STAGE_META[stageKey];
    const complete = finalStageIsComplete(rows, meta.tableCount);
    const qualifiers = complete ? getFinalQualifiers(stageKey, rows) : [];
    return {
      key: stageKey,
      ...meta,
      source,
      rows,
      started: finalStageIsStarted(rows),
      complete,
      qualifiedPlayerIds: qualifiers.map((row) => row.playerId),
    };
  }

  function getFinalsData() {
    const apiFinals = state.data?.finals || {};
    const eliminationComplete = typeof apiFinals.eliminationComplete === "boolean"
      ? apiFinals.eliminationComplete
      : eliminationsAreComplete();
    const standings = eliminationComplete ? computeStandings(3) : [];
    const seedByPlayer = new Map(standings.map((entry, index) => [entry.id, index + 1]));

    const apiRows = (stageKey) => {
      const rows = apiFinals?.[stageKey]?.rows;
      return Array.isArray(rows)
        ? rows.map((row) => normaliseFinalRow(row, seedByPlayer)).filter((row) => row.table && row.playerId)
        : [];
    };

    let ottaviRows = apiRows("ottavi");
    let ottaviSource = apiFinals?.ottavi?.source || "waiting";
    if (!ottaviRows.length && eliminationComplete) {
      ottaviRows = buildOttaviRows(standings);
      ottaviSource = "preview";
    }
    const ottavi = makeFinalStage("ottavi", ottaviRows, ottaviSource);

    let quartiRows = apiRows("quarti");
    let quartiSource = apiFinals?.quarti?.source || "waiting";
    if (!quartiRows.length && ottavi.complete) {
      quartiRows = buildQuartiRows(ottavi.rows);
      quartiSource = "preview";
    }
    const quarti = makeFinalStage("quarti", quartiRows, quartiSource);

    let semifinaliRows = apiRows("semifinali");
    let semifinaliSource = apiFinals?.semifinali?.source || "waiting";
    if (!semifinaliRows.length && quarti.complete) {
      semifinaliRows = buildSemifinaliRows(quarti.rows);
      semifinaliSource = "preview";
    }
    const semifinali = makeFinalStage("semifinali", semifinaliRows, semifinaliSource);

    let finaleRows = apiRows("finale");
    let finaleSource = apiFinals?.finale?.source || "waiting";
    if (!finaleRows.length && semifinali.complete) {
      finaleRows = buildFinaleRows(semifinali.rows);
      finaleSource = "preview";
    }
    const finale = makeFinalStage("finale", finaleRows, finaleSource);

    let currentStage = "eliminatorie";
    if (eliminationComplete) currentStage = "ottavi";
    if (ottavi.complete) currentStage = "quarti";
    if (quarti.complete) currentStage = "semifinali";
    if (semifinali.complete) currentStage = "finale";
    if (finale.complete) currentStage = "concluso";

    const winnerRow = finale.complete ? finale.rows.find((row) => row.position === 1) : null;
    return {
      eliminationComplete,
      cutoffTie: Boolean(apiFinals.cutoffTie),
      currentStage,
      winnerPlayerId: winnerRow?.playerId || apiFinals.winnerPlayerId || null,
      ottavi,
      quarti,
      semifinali,
      finale,
    };
  }

  function finalTableGroups(stage) {
    const groups = groupFinalRows(stage?.rows || []);
    return [...groups.entries()]
      .sort(([a], [b]) => a.localeCompare(b, "it"))
      .map(([table, rows]) => ({
        table,
        rows,
        completed: finalTableIsComplete(rows),
        started: finalStageIsStarted(rows),
      }));
  }

  function finalProgressStats(finals) {
    const stages = FINAL_STAGE_ORDER.map((key) => finals[key]);
    const total = stages.reduce((sum, stage) => sum + stage.tableCount, 0);
    const groups = stages.flatMap(finalTableGroups);
    return {
      total,
      completed: groups.filter((group) => group.completed).length,
      live: groups.filter((group) => group.started && !group.completed).length,
      todo: total - groups.filter((group) => group.completed || group.started).length,
    };
  }

  function finalStageStatus(stage) {
    if (stage.complete) return { label: "Concluso", className: "done" };
    if (stage.started) return { label: "In corso", className: "live" };
    if (stage.rows.length) return { label: "Pronto", className: "" };
    return { label: "Da definire", className: "" };
  }

  function currentFinalStageLabel(finals) {
    if (finals.currentStage === "concluso") return "Torneo concluso";
    if (FINAL_STAGE_META[finals.currentStage]) return FINAL_STAGE_META[finals.currentStage].label;
    return "Fase eliminatoria conclusa";
  }

  function renderFinalsBoard() {
    const finals = getFinalsData();
    const playerMap = buildPlayerMap(state.data);
    const board = $("#finals-board");
    if (!board) return;

    board.innerHTML = FINAL_STAGE_ORDER.map((stageKey) => {
      const stage = finals[stageKey];
      const status = finalStageStatus(stage);
      const current = finals.currentStage === stageKey;
      const qualifiedIds = new Set(stage.qualifiedPlayerIds || []);
      const groups = finalTableGroups(stage);

      const tablesMarkup = groups.length
        ? groups.map((group) => `
          <div class="bracket-table">
            <span>Tavolo ${escapeHtml(group.table)}</span>
            ${group.rows.map((row) => {
              const player = playerMap.get(row.playerId);
              const advanced = qualifiedIds.has(row.playerId);
              return `<div class="bracket-player ${advanced ? "advanced" : ""}"><i>${row.seed || "—"}</i><b>${escapeHtml(player?.name || `Giocatore ${row.playerId}`)}</b></div>`;
            }).join("")}
          </div>`).join("")
        : '<div class="bracket-table"><div class="bracket-player placeholder">Partecipanti da definire</div></div>';

      const winner = stageKey === "finale" && finals.winnerPlayerId
        ? `<div class="bracket-winner"><div><small>Vincitore del torneo</small><strong>${escapeHtml(playerMap.get(finals.winnerPlayerId)?.name || `Giocatore ${finals.winnerPlayerId}`)}</strong></div></div>`
        : tablesMarkup;

      return `
        <section class="bracket-stage ${current ? "current" : ""} ${stage.complete ? "complete" : ""}">
          <div class="bracket-stage-header">
            <div><strong>${stage.label}</strong><small>${stage.playerCount} giocatori · ${stage.tableCount} ${stage.tableCount === 1 ? "tavolo" : "tavoli"}</small></div>
            <span class="bracket-status ${status.className}">${status.label}</span>
          </div>
          ${winner}
        </section>`;
    }).join("");
  }

  function renderFinals() {
    const finals = getFinalsData();
    if (!FINAL_STAGE_META[state.finalStage]) state.finalStage = finals.currentStage === "concluso" ? "finale" : (FINAL_STAGE_META[finals.currentStage] ? finals.currentStage : "ottavi");
    const stage = finals[state.finalStage];
    const playerMap = buildPlayerMap(state.data);
    const qualifiedIds = new Set(stage.qualifiedPlayerIds || []);
    const groups = finalTableGroups(stage);
    const complete = groups.filter((group) => group.completed).length;
    const live = groups.filter((group) => group.started && !group.completed).length;
    const todo = stage.tableCount - complete - live;

    $("#finals-status").textContent = currentFinalStageLabel(finals);
    $$(".final-round-tab").forEach((button) => {
      const key = button.dataset.finalStage;
      button.classList.toggle("active", key === state.finalStage);
      button.classList.toggle("current", key === finals.currentStage);
    });

    $("#final-tables-summary").textContent = `${stage.tableCount} ${stage.tableCount === 1 ? "tavolo" : "tavoli"} · ${complete} completati · ${live} in corso · ${Math.max(0, todo)} da giocare`;
    const grid = $("#final-tables-grid");
    if (!groups.length) {
      grid.innerHTML = '<div class="panel empty-state">Partecipanti ancora da definire.</div>';
      return;
    }

    grid.innerHTML = groups.map((group) => {
      const orderedRows = group.completed ? [...group.rows].sort((a, b) => a.position - b.position) : group.rows;
      const statusClass = group.completed ? "done" : group.started ? "live" : "";
      const statusLabel = group.completed ? "Concluso" : group.started ? "In corso" : "Da giocare";
      return `
        <article class="table-card">
          <header class="table-card-header"><h3>Tavolo ${escapeHtml(group.table)}</h3><span class="table-status ${statusClass}">${statusLabel}</span></header>
          <div class="table-players">
            ${orderedRows.map((row) => {
              const player = playerMap.get(row.playerId);
              const completed = isCompletedRow(row);
              const hasPoints = row.gamePoints !== null;
              const advanced = qualifiedIds.has(row.playerId);
              const positionBadge = completed
                ? `<span class="seat-or-place">${row.position}°</span>`
                : "";
              return `
                <div class="table-player final-player ${completed ? "" : "final-player-pending"} ${row.position === 1 ? "winner" : ""} ${advanced ? "advanced" : ""}">
                  ${positionBadge}
                  <div class="table-player-name">
                    <strong>${escapeHtml(player?.name || `Giocatore ${row.playerId}`)}</strong>
                    ${advanced ? '<span class="advance-badge">Qualificato</span>' : ""}
                  </div>
                  <div class="table-player-result"><strong>${hasPoints ? row.gamePoints : "—"}</strong><small>punti partita</small></div>
                </div>`;
            }).join("")}
          </div>
        </article>`;
    }).join("");
  }

  function renderHeader() {
    const { config, source, appsScriptError, excelError } = state.data;
    $("#tournament-title").textContent = config.title;
    const finals = getFinalsData();
    $("#phase-label").textContent = finals.eliminationComplete ? currentFinalStageLabel(finals) : config.phase;

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
    const eliminationStats = globalTableStats();
    const stage = currentStage();
    const currentRound = currentResultsRound();
    const finals = getFinalsData();
    const finalStats = finalProgressStats(finals);
    const usingFinals = finals.eliminationComplete;
    const stats = usingFinals ? finalStats : eliminationStats;
    const percent = stats.total ? Math.round((stats.completed / stats.total) * 100) : 0;

    $("#stat-players").textContent = config.playersCount;
    $("#stat-cutoff").textContent = usingFinals ? `${config.cutoff} qualificati` : `Top ${config.cutoff}`;
    $("#stat-cutoff-label").textContent = usingFinals ? "Fase finale" : "Qualificazione";
    $("#stat-cutoff-detail").textContent = usingFinals ? "accesso ufficiale agli ottavi" : "accedono agli ottavi";
    $("#round-status").textContent = stage.label;

    $("#progress-title").textContent = usingFinals ? "Fasi finali" : "Fase eliminatoria";
    $("#progress-ring").style.setProperty("--progress", percent);
    $("#progress-percent").textContent = `${percent}%`;
    $("#progress-todo").textContent = stats.todo;
    $("#progress-live").textContent = stats.live;
    $("#progress-done").textContent = stats.completed;

    const rankingRound = usingFinals ? 3 : (currentRound || 1);
    const top = computeStandings(rankingRound).filter((entry) => entry.played > 0).slice(0, usingFinals ? config.cutoff : 10);
    const container = $("#top-ranking");
    $("#ranking-panel-kicker").textContent = usingFinals ? "Classifica eliminatoria ufficiale" : "Classifica provvisoria";
    $("#ranking-panel-title").textContent = usingFinals ? "Qualificati agli ottavi" : "Prime posizioni";

    if (!top.length) {
      container.className = "compact-ranking empty-state";
      container.textContent = "Nessun tavolo ancora completato.";
    } else {
      container.className = "compact-ranking";
      container.innerHTML = top.slice(0, 10).map((entry) => `
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

    renderFinalsBoard();
  }

  function renderRanking() {
    const round = currentResultsRound() || 1;
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
          <td class="center">${entry.sole}</td>
          <td class="center score-main">${entry.points}</td>
          <td class="center tie-col">${displayMetric(entry.bestPosition)}</td>
          <td class="center tie-col">${displayMetric(entry.minGamePoints)}</td>
          <td class="center tie-col">${displayMetric(entry.totalGamePoints)}</td>
        </tr>`;
    }).join("");

    if (!filtered.length) {
      body.innerHTML = '<tr><td colspan="8" class="center">Nessun giocatore trovato.</td></tr>';
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

  function closeEliminationMenu() {
    const dropdown = $(".nav-dropdown");
    const toggle = $(".nav-dropdown-toggle");
    if (!dropdown || !toggle) return;
    dropdown.classList.remove("open");
    toggle.setAttribute("aria-expanded", "false");
  }

  function setView(view) {
    state.activeView = view;
    $$(".view").forEach((section) => section.classList.toggle("active", section.id === view));

    $$(".nav-button[data-view]").forEach((button) => {
      const active = button.dataset.view === view;
      button.classList.toggle("active", active);
      if (active) button.setAttribute("aria-current", "page");
      else button.removeAttribute("aria-current");
    });

    const eliminationActive = view === "classifica" || view === "tavoli";
    const eliminationToggle = $(".nav-dropdown-toggle");
    if (eliminationToggle) {
      eliminationToggle.classList.toggle("active", eliminationActive);
      if (eliminationActive) eliminationToggle.setAttribute("aria-current", "page");
      else eliminationToggle.removeAttribute("aria-current");
    }

    $$(".nav-subbutton").forEach((button) => {
      const active = button.dataset.view === view;
      button.classList.toggle("active", active);
      if (active) button.setAttribute("aria-current", "page");
      else button.removeAttribute("aria-current");
    });

    closeEliminationMenu();
    history.replaceState(null, "", `#${view}`);
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function bindEvents() {
    $$(".nav-button[data-view]").forEach((button) => button.addEventListener("click", () => setView(button.dataset.view)));
    $$(".nav-subbutton[data-view]").forEach((button) => button.addEventListener("click", () => setView(button.dataset.view)));

    const eliminationDropdown = $(".nav-dropdown");
    const eliminationToggle = $(".nav-dropdown-toggle");
    if (eliminationDropdown && eliminationToggle) {
      eliminationToggle.addEventListener("click", (event) => {
        event.stopPropagation();
        const willOpen = !eliminationDropdown.classList.contains("open");
        eliminationDropdown.classList.toggle("open", willOpen);
        eliminationToggle.setAttribute("aria-expanded", String(willOpen));
      });

      document.addEventListener("click", (event) => {
        if (!eliminationDropdown.contains(event.target)) closeEliminationMenu();
      });

      document.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
          closeEliminationMenu();
          eliminationToggle.focus();
        }
      });
    }

    $$("[data-open-view]").forEach((button) => button.addEventListener("click", () => setView(button.dataset.openView)));
    $("#ranking-search").addEventListener("input", renderRanking);

    $$(".round-tab").forEach((button) => button.addEventListener("click", () => {
      state.tableRound = Number(button.dataset.round);
      $$(".round-tab").forEach((item) => item.classList.toggle("active", item === button));
      renderTables();
    }));

    $$(".final-round-tab").forEach((button) => button.addEventListener("click", () => {
      state.finalStage = button.dataset.finalStage;
      renderFinals();
    }));
  }

  function renderAll() {
    renderHeader();
    renderDashboard();
    renderRanking();
    renderTables();
    renderFinals();
  }


  async function refreshFromSource() {
    if (state.refreshInFlight) return;
    state.refreshInFlight = true;

    try {
      const freshData = await loadTournamentData({ allowFallback: false });
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
      const initialFinals = getFinalsData();
      state.finalStage = initialFinals.currentStage === "concluso"
        ? "finale"
        : (FINAL_STAGE_META[initialFinals.currentStage] ? initialFinals.currentStage : "ottavi");
      $$(".round-tab").forEach((button) => button.classList.toggle("active", Number(button.dataset.round) === state.tableRound));
      renderAll();
      startAutoRefresh();

      const hashView = location.hash.replace("#", "");
      if (["dashboard", "classifica", "tavoli", "finali", "regolamento"].includes(hashView)) setView(hashView);
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
