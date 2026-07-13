const SPREADSHEET_ID = 'INCOLLA_QUI_ID_GOOGLE_SHEET';

function doGet(e) {
  try {
    const data = buildTournamentData_();
    const json = JSON.stringify(data);
    const callback = String((e && e.parameter && (e.parameter.callback || e.parameter.prefix)) || '');

    if (callback) {
      if (!/^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(callback)) {
        throw new Error('Nome callback non valido.');
      }
      return ContentService
        .createTextOutput(`${callback}(${json});`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const payload = JSON.stringify({
      error: true,
      message: error && error.message ? error.message : String(error)
    });

    return ContentService
      .createTextOutput(payload)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function testConnection() {
  const data = buildTournamentData_();
  console.log(JSON.stringify(data, null, 2));
}

function buildTournamentData_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const configValues = spreadsheet.getSheetByName('Config').getDataRange().getValues();
  const configMap = {};

  configValues.forEach(row => {
    const key = normalize_(row[0]);
    if (key) configMap[key] = row[1];
  });

  const playersCount = toNumber_(configMap['numero giocatori']) || 65;
  const cutoff = toNumber_(configMap['posti qualificazione']) || 25;

  const players = readPlayers_(spreadsheet, playersCount);
  const results = readResults_(spreadsheet, playersCount);

  return {
    source: 'apps-script',
    generatedAt: new Date().toISOString(),
    config: {
      title: String(configMap['nome torneo'] || "Torneo Ass 'e Mazz 2026"),
      playersCount: playersCount,
      cutoff: cutoff,
      phase: String(configMap['fase'] || 'Eliminatorie')
    },
    players: players,
    results: results
  };
}

function readPlayers_(spreadsheet, playersCount) {
  const sheet = spreadsheet.getSheetByName('Giocatori');
  if (!sheet) throw new Error('Foglio Giocatori non trovato.');

  const values = sheet.getDataRange().getValues();
  const headerIndex = findHeaderRow_(values, ['id', 'nome e cognome']);
  const headers = values[headerIndex].map(normalize_);

  return values.slice(headerIndex + 1)
    .map(row => rowToObject_(headers, row))
    .map(row => {
      const id = toNumber_(row['id']);
      if (!id || id > playersCount) return null;
      return {
        id: id,
        name: String(row['nome e cognome'] || '').trim(),
        nickname: String(row['soprannome (facoltativo)'] || '').trim()
      };
    })
    .filter(Boolean);
}

function readResults_(spreadsheet, playersCount) {
  const sheetName = `Risultati_${playersCount}`;
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Foglio ${sheetName} non trovato.`);

  const values = sheet.getDataRange().getValues();
  const headerIndex = findHeaderRow_(values, ['turno', 'tavolo', 'giocatore id', 'posizione']);
  const headers = values[headerIndex].map(normalize_);

  return values.slice(headerIndex + 1)
    .map(row => rowToObject_(headers, row))
    .map(row => ({
      turn: toNumber_(row['turno']),
      table: String(row['tavolo'] || '').trim().toUpperCase(),
      seat: toNumber_(row['sedia']),
      playerId: toNumber_(row['giocatore id']),
      gamePoints: toNumber_(row['punti partita']),
      position: toNumber_(row['posizione']),
      sole: toNumber_(row['sole']) || 0,
      note: String(row['note'] || '').trim()
    }))
    .filter(row => row.turn && row.table && row.playerId);
}

function findHeaderRow_(values, requiredHeaders) {
  const index = values.findIndex(row => {
    const normalized = row.map(normalize_);
    return requiredHeaders.every(header => normalized.includes(header));
  });

  if (index === -1) {
    throw new Error(`Intestazione non trovata: ${requiredHeaders.join(', ')}`);
  }
  return index;
}

function rowToObject_(headers, row) {
  const object = {};
  headers.forEach((header, index) => {
    if (header) object[header] = row[index];
  });
  return object;
}

function normalize_(value) {
  return String(value == null ? '' : value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function toNumber_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const number = Number(String(value).replace(',', '.'));
  return Number.isFinite(number) ? number : null;
}
