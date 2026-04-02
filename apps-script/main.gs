const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';
const SHEET_NAME = 'YOUR_PRODUCT_SHEET_NAME';
const SHEET_PROGRESS = 'YOUR_PROGRESS_SHEET_NAME';
const API_KEY_SUPPLIER = 'YOUR_SUPPLIER_API_KEY';
const LANG = 'en';
const CURRENCY = 'EUR';
const MAX_RUNTIME_MS = 300000;
const CACHE_KEY = 'YOUR_CACHE_KEY';
const CACHE_TTL = 21600;
const CHUNK_SIZE = 150;

function getHeaders() {
  return { Authorization: 'Token ' + API_KEY_SUPPLIER };
}

const HEADERS = [
  'ID', 'Titolo', 'Descrizione', 'Prezzo', 'Valuta',
  'Immagine', 'Link Checkout', 'Rating', 'Recensioni',
  'Città', 'Paese', 'Venue', 'Aggiornato il', 'URL Prodotto'
];

function mapProdotto(p) {
  return [
    p.id?.toString() || '',
    p.title || '',
    p.summary || '',
    p.price || '',
    p.currency || CURRENCY,
    '',
    p.product_checkout_url || '',
    p.ratings?.average || '',
    p.ratings?.total || '',
    p.city_name || '',
    p.country_name || '',
    p.venue?.name || '',
    new Date().toLocaleDateString('it-IT'),
    p.product_url || ''
  ];
}

function getProgress() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_PROGRESS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROGRESS);
    sheet.getRange('A1:C1').setValues([['last_page', 'total_pages', 'totale_prodotti']]);
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2 || !data[1][0]) return null;

  return {
    lastPage: Number(data[1][0]),
    totalPages: Number(data[1][1]),
    totale: Number(data[1][2])
  };
}

function saveProgress(lastPage, totalPages, totale) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ss.getSheetByName(SHEET_PROGRESS).getRange('A2:C2').setValues([[lastPage, totalPages, totale]]);
}

function clearProgress() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PROGRESS);
  if (sheet) sheet.getRange('A2:C2').clearContent();
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'continuaScaricaCatalogo')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function deleteImgTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'continuaArricchisciImmagini')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function scaricaCatalogoCompleto() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 14).setValues([HEADERS]).setFontWeight('bold');

  saveProgress(1, 0, 0);
  deleteTriggers();
  ScriptApp.newTrigger('continuaScaricaCatalogo').timeBased().everyMinutes(1).create();
  continuaScaricaCatalogo();
}

function continuaScaricaCatalogo() {
  const progress = getProgress();
  if (!progress) {
    deleteTriggers();
    return;
  }

  let { lastPage, totalPages, totale } = progress;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const startTime = Date.now();

  if (!totalPages || totalPages === 0) {
    const firstRes = UrlFetchApp.fetch(
      `https://api.supplier.com/v2/products?lang=${LANG}&currency=${CURRENCY}&page_size=100&page=1`,
      { headers: getHeaders(), muteHttpExceptions: true }
    );

    if (firstRes.getResponseCode() !== 200) {
      Logger.log(`Errore bootstrap catalogo: HTTP ${firstRes.getResponseCode()}`);
      return;
    }

    const firstJson = JSON.parse(firstRes.getContentText());
    totalPages = Math.ceil((firstJson.pagination?.total || 0) / 100);
    Logger.log(`📦 Catalogo: ${firstJson.pagination?.total || 0} prodotti in ${totalPages} pagine`);
  }

  let page = lastPage;

  while (page <= totalPages) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      saveProgress(page, totalPages, totale);
      Logger.log(`⏸ Pausa a pagina ${page}/${totalPages} — salvati: ${totale}`);
      return;
    }

    const url = `https://api.supplier.com/v2/products?lang=${LANG}&currency=${CURRENCY}&page_size=100&page=${page}`;
    const res = UrlFetchApp.fetch(url, { headers: getHeaders(), muteHttpExceptions: true });

    if (res.getResponseCode() !== 200) {
      Logger.log(`Errore HTTP ${res.getResponseCode()} a pagina ${page} — salto`);
      page++;
      continue;
    }

    const prodotti = JSON.parse(res.getContentText()).products || [];
    const rows = prodotti.map(mapProdotto);

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 14).setValues(rows);
      totale += rows.length;
    }

    Logger.log(`✅ Pagina ${page}/${totalPages} — totale salvati: ${totale}`);
    saveProgress(page + 1, totalPages, totale);
    page++;
    Utilities.sleep(200);
  }

  clearProgress();
  deleteTriggers();
  Logger.log(`🎉 Completato! ${totale} prodotti scaricati.`);
  ss.toast(`🎉 Catalogo completo: ${totale} prodotti!`, 'Download completato', 10);
}

function avviaArricchisciImmagini() {
  deleteImgTriggers();
  ScriptApp.newTrigger('continuaArricchisciImmagini').timeBased().everyMinutes(1).create();
  continuaArricchisciImmagini();
}

function continuaArricchisciImmagini() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const startTime = Date.now();
  let aggiornati = 0;
  let completato = true;

  for (let i = 1; i < values.length; i++) {
    if (Date.now() - startTime > MAX_RUNTIME_MS) {
      Logger.log(`⏸ Pausa — riprendo da riga ${i + 1}. Aggiornati: ${aggiornati}`);
      completato = false;
      return;
    }

    const imgCol = values[i][5];
    const urlProdotto = values[i][13];
    if (imgCol || !urlProdotto) continue;

    try {
      const res = UrlFetchApp.fetch(urlProdotto, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) continue;

      const html = res.getContentText();
      const match = html.match(/https:\/\/aws-supplier-cdn\.imgix\.net\/images\/content\/[a-f0-9]+\.jpg/i);

      if (match && match[0]) {
        const imgUrl = match[0] + '?auto=format,compress&fit=crop&w=600&h=400&q=70';
        sheet.getRange(i + 1, 6).setValue(imgUrl);
        aggiornati++;
        Logger.log(`✅ Riga ${i + 1}: ${imgUrl}`);
      }

      Utilities.sleep(300);
    } catch (e) {
      Logger.log(`Errore riga ${i + 1}: ${e}`);
    }
  }

  if (completato) {
    deleteImgTriggers();
    Logger.log(`🎉 Immagini completate: ${aggiornati} aggiornate.`);
    ss.toast(`🎉 Immagini complete: ${aggiornati} aggiornate!`, 'Arricchimento completato', 10);
  }
}

function readFromSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  rows.shift();

  return rows
    .filter(r => r[0])
    .map(r => ({
      id: r[0],
      titolo: r[1],
      descrizione: r[2],
      prezzo: r[3],
      valuta: r[4],
      immagine: r[5],
      link: r[6],
      rating: r[7],
      reviews: r[8],
      citta: r[9],
      paese: r[10],
      venue: r[11],
      url_prodotto: r[13]
    }));
}

function warmUpCache() {
  const cache = CacheService.getScriptCache();
  const allData = readFromSheet();

  const slim = allData.map(p => {
    const o = {};
    if (p.id) o.id = p.id;
    if (p.titolo) o.t = String(p.titolo);
    if (p.descrizione) o.d = String(p.descrizione).substring(0, 120);
    if (p.prezzo) o.p = p.prezzo;
    if (p.immagine) o.i = String(p.immagine);
    if (p.link) o.l = String(p.link);
    if (p.rating) o.r = p.rating;
    if (p.reviews) o.rv = p.reviews;
    if (p.citta) o.c = String(p.citta);
    if (p.paese) o.pa = String(p.paese);
    if (p.venue) o.v = String(p.venue);
    return o;
  });

  const chunks = [];
  for (let i = 0; i < slim.length; i += CHUNK_SIZE) {
    chunks.push(slim.slice(i, i + CHUNK_SIZE));
  }

  const pairs = {};
  chunks.forEach((chunk, index) => {
    pairs[`${CACHE_KEY}_${index}`] = JSON.stringify(chunk);
  });
  pairs[`${CACHE_KEY}_count`] = String(chunks.length);

  cache.putAll(pairs, CACHE_TTL);
  Logger.log(`✅ Cache warm: ${slim.length} prodotti in ${chunks.length} chunk`);
}

function readFromCache() {
  const cache = CacheService.getScriptCache();
  const countStr = cache.get(`${CACHE_KEY}_count`);
  if (!countStr) return null;

  const count = parseInt(countStr, 10);
  const keys = Array.from({ length: count }, (_, i) => `${CACHE_KEY}_${i}`);
  const chunks = cache.getAll(keys);
  if (Object.keys(chunks).length < count) return null;

  let allData = [];
  for (let i = 0; i < count; i++) {
    const chunk = chunks[`${CACHE_KEY}_${i}`];
    if (!chunk) return null;

    const parsed = JSON.parse(chunk).map(o => ({
      id: o.id,
      titolo: o.t,
      descrizione: o.d,
      prezzo: o.p,
      immagine: o.i,
      link: o.l,
      rating: o.r,
      reviews: o.rv,
      citta: o.c,
      paese: o.pa,
      venue: o.v
    }));

    allData = allData.concat(parsed);
  }

  return allData;
}

function invalidateCache() {
  const cache = CacheService.getScriptCache();
  const countStr = cache.get(`${CACHE_KEY}_count`);
  if (!countStr) {
    Logger.log('ℹ️ Cache già vuota.');
    return;
  }

  const count = parseInt(countStr, 10);
  const keys = Array.from({ length: count }, (_, i) => `${CACHE_KEY}_${i}`);
  keys.push(`${CACHE_KEY}_count`);
  cache.removeAll(keys);
  Logger.log('✅ Cache invalidata.');
}

function doGet(e) {
  let allData = readFromCache();

  if (!allData) {
    allData = readFromSheet();
    warmUpCache();
  }

  const filterCity = e?.parameter?.city?.toLowerCase() || null;
  const filterCountry = e?.parameter?.country?.toLowerCase() || null;
  const filterQ = e?.parameter?.q?.toLowerCase() || null;
  const pageSize = parseInt(e?.parameter?.page_size, 10) || 100;
  const page = parseInt(e?.parameter?.page, 10) || 1;

  let data = allData;
  if (filterCity) data = data.filter(r => r.citta?.toLowerCase().includes(filterCity));
  if (filterCountry) data = data.filter(r => r.paese?.toLowerCase().includes(filterCountry));
  if (filterQ) data = data.filter(r => r.titolo?.toLowerCase().includes(filterQ));

  const total = data.length;
  const paginated = data.slice((page - 1) * pageSize, page * pageSize);

  const cleanProducts = paginated.map(p => ({
    Titolo: String(p.titolo || ''),
    Descrizione: String(p.descrizione || ''),
    Prezzo: p.prezzo || 0,
    Immagine: String(p.immagine || ''),
    'Link Checkout': String(p.link || ''),
    Rating: p.rating || '',
    Recensioni: p.reviews || '',
    City: String(p.citta || ''),
    Country: String(p.paese || ''),
    Venue: String(p.venue || '')
  }));

  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      cached: true,
      total,
      page,
      page_size: pageSize,
      products: cleanProducts
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function testDieciProdotti() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 14).setValues([HEADERS]).setFontWeight('bold');

  const url = `https://api.supplier.com/v2/products?lang=${LANG}&currency=${CURRENCY}&page_size=10&page=1`;
  const res = UrlFetchApp.fetch(url, { headers: getHeaders(), muteHttpExceptions: true });
  const prodotti = JSON.parse(res.getContentText()).products || [];
  const rows = prodotti.map(mapProdotto);

  sheet.getRange(2, 1, rows.length, 14).setValues(rows);
  Logger.log(`✅ Test completato: ${rows.length} prodotti scritti.`);
}

function testImmagini3Righe() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();

  let testati = 0;
  for (let i = 1; i < values.length && testati < 3; i++) {
    const urlProdotto = values[i][13];
    if (!urlProdotto) continue;

    const res = UrlFetchApp.fetch(urlProdotto, { muteHttpExceptions: true });
    const html = res.getContentText();
    const match = html.match(/https:\/\/aws-supplier-cdn\.imgix\.net\/images\/content\/[a-f0-9]+\.jpg/i);

    Logger.log(`Riga ${i + 1} | immagine: ${match ? match[0] + '?auto=format,compress&fit=crop&w=600&h=400&q=70' : 'NON TROVATO'}`);
    testati++;
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎟 Supplier')
    .addItem('📦 Scarica catalogo completo', 'scaricaCatalogoCompleto')
    .addItem('🖼 Arricchisci immagini', 'avviaArricchisciImmagini')
    .addSeparator()
    .addItem('🧪 Test 10 prodotti', 'testDieciProdotti')
    .addItem('🧪 Test immagini 3 righe', 'testImmagini3Righe')
    .addSeparator()
    .addItem('♻️ Invalida cache', 'invalidateCache')
    .addToUi();
}
