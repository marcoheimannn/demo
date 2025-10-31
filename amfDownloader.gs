const YEAR = 2024;
const TIMEZONE = 'Europe/Paris';
const RATE_LIMIT_PER_MINUTE = 60;
const INTER_PAGE_DELAY_MS = 200;
const MAX_RETRIES = 3;
const BASE_URL = 'https://www.info-financiere.gouv.fr/api/explore/v2.1';
const DATASET_ID = 'flux-amf-new-prod';
const RECORDS_ENDPOINT = BASE_URL + '/catalog/datasets/' + DATASET_ID + '/records';
const CSV_ENDPOINT = BASE_URL + '/catalog/datasets/' + DATASET_ID + '/exports/csv';
const DATASET_ENDPOINT = BASE_URL + '/catalog/datasets/' + DATASET_ID;
const TARGET_DRIVE_FOLDER_ID = '1gFgFdAxruSwmQcFyxPkyk5j8cRFI-4NL';
const SPREADSHEET_ID = '[SPREADSHEET_ID]';
const STATE_PROPERTY_KEY = 'AMF_STATE_MAP';
const LAST_RUN_PROPERTY_KEY = 'AMF_LAST_RUN_ISO';
const DATASET_FIELDS_PROPERTY_KEY = 'AMF_DATASET_FIELDS';
const RATE_INTERVAL_MS = Math.ceil(60000 / RATE_LIMIT_PER_MINUTE);
const DOWNLOAD_FIELD_REGEX = /(url|lien|telechargement|fichier|href)/i;
const URL_REGEX = /^https?:\/\//i;

var rateLimiter = (function () {
  var lastRequestTime = 0;
  return {
    wait: function () {
      var now = Date.now();
      var waitMs = lastRequestTime + RATE_INTERVAL_MS - now;
      if (waitMs > 0) {
        Utilities.sleep(waitMs);
      }
      lastRequestTime = Date.now();
    }
  };
})();

function run() {
  processIssuers({ dryRun: false });
}

function testOnce() {
  processIssuers({ dryRun: true, maxIssuers: 2, maxPages: 1 });
}

function createTrigger() {
  var existing = ScriptApp.getProjectTriggers();
  existing.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'run') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('run')
    .timeBased()
    .atHour(2)
    .nearMinute(15)
    .everyDays(1)
    .inTimezone(TIMEZONE)
    .create();
}

function processIssuers(options) {
  options = options || {};
  var dryRun = options.dryRun === true;
  var maxIssuers = options.maxIssuers || null;
  var maxPages = options.maxPages || null;
  var props = PropertiesService.getScriptProperties();
  var stateMap = loadStateMap(props);
  var apiKey = props.getProperty('APIKEY');
  var datasetFields = loadDatasetFields(props, apiKey);
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('CAC40');
  if (!sheet) {
    throw new Error('Sheet CAC40 not found');
  }
  var rows = sheet.getDataRange().getValues();
  if (rows.length === 0) {
    return;
  }
  var header = rows.shift();
  var colIndex = buildColumnIndex(header);
  var issuerCount = 0;
  rows.forEach(function (row) {
    if (maxIssuers !== null && issuerCount >= maxIssuers) {
      return;
    }
    var legalName = getCellValue(row, colIndex, 'LegalName');
    var keyword = getCellValue(row, colIndex, 'Keyword');
    var isin = getCellValue(row, colIndex, 'ISIN');
    var lei = getCellValue(row, colIndex, 'LEI');
    if (!isin || !legalName) {
      return;
    }
    issuerCount++;
    var issuerOptions = {
      legalName: legalName,
      keyword: keyword,
      isin: isin,
      lei: lei,
      dryRun: dryRun,
      maxPages: maxPages,
      apiKey: apiKey,
      datasetFields: datasetFields,
      stateMap: stateMap,
      props: props
    };
    processIssuer(issuerOptions);
  });
  if (!dryRun) {
    saveStateMap(props, stateMap);
    props.setProperty(LAST_RUN_PROPERTY_KEY, new Date().toISOString());
  }
}

function buildColumnIndex(header) {
  var index = {};
  header.forEach(function (name, idx) {
    if (name) {
      index[String(name).trim()] = idx;
    }
  });
  return index;
}

function getCellValue(row, indexMap, key) {
  var idx = indexMap[key];
  if (idx === undefined) {
    return '';
  }
  var value = row[idx];
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function loadStateMap(props) {
  var raw = props.getProperty(STATE_PROPERTY_KEY);
  if (!raw) {
    return {};
  }
  try {
    return JSON.parse(raw);
  } catch (err) {
    Logger.log('Failed to parse state map, resetting: ' + err);
    return {};
  }
}

function saveStateMap(props, stateMap) {
  props.setProperty(STATE_PROPERTY_KEY, JSON.stringify(stateMap));
}

function loadDatasetFields(props, apiKey) {
  var cached = props.getProperty(DATASET_FIELDS_PROPERTY_KEY);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      Logger.log('Failed to parse cached dataset fields: ' + e);
    }
  }
  var response = fetchJson(DATASET_ENDPOINT, { apikey: apiKey });
  if (!response || !response.dataset || !response.dataset.fields) {
    return [];
  }
  var fields = response.dataset.fields
    .map(function (field) {
      return field.name;
    })
    .filter(function (name) {
      return !!name;
    });
  props.setProperty(DATASET_FIELDS_PROPERTY_KEY, JSON.stringify(fields));
  return fields;
}

function processIssuer(options) {
  var start = Date.now();
  var records = fetchIssuerRecords(options);
  var savedFiles = [];
  var stateMap = options.stateMap;
  var dryRun = options.dryRun;
  records.forEach(function (record) {
    var recordId = String(record.id);
    if (!recordId) {
      return;
    }
    var metaJson = JSON.stringify(record);
    var digest = toHex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, metaJson));
    if (stateMap[recordId] === digest && !dryRun) {
      return;
    }
    var download = pickBestDownload(record);
    if (!download) {
      Logger.log('No download URL for record ' + recordId + ' (' + options.isin + ')');
      if (!dryRun) {
        stateMap[recordId] = digest;
      }
      return;
    }
    var dateStamp = formatDateForRecord(record.uin_dat_amf);
    var issuerSlug = makeSlug(record.emetteur_denomination || options.legalName || options.isin);
    var amendmentSuffix = buildAmendmentSuffix(record.libelles);
    var baseName = options.isin + '-URD-' + dateStamp + '-' + issuerSlug + '-' + recordId + amendmentSuffix;
    var extension = determineExtension(download.url, download.contentType);
    var targetFolderPath = [options.isin, String(YEAR), dateStamp];
    if (dryRun) {
      savedFiles.push(baseName + '.' + extension);
      return;
    }
    var folder = ensureFolderPath(TARGET_DRIVE_FOLDER_ID, targetFolderPath);
    var blob = fetchFile(download.url, options.apiKey, extension);
    if (download.contentType) {
      blob.setContentType(download.contentType);
    }
    blob.setName(baseName + '.' + extension);
    var file = folder.createFile(blob);
    savedFiles.push(file.getName());
    if (extension.toLowerCase() === 'zip') {
      handleZipExtraction(file, folder);
    }
    writeMetadataFile(folder, baseName, metaJson);
    stateMap[recordId] = digest;
  });
  var ms = Date.now() - start;
  logIssuer(options.isin, records.length, savedFiles, ms);
}

function fetchIssuerRecords(options) {
  var whereClause = buildWhereClause(options.isin);
  var selectFields = buildSelectFields(options.datasetFields);
  var params = {
    where: whereClause,
    order_by: '-uin_dat_amf',
    limit: 100,
    offset: 0,
    select: selectFields.join(','),
    apikey: options.apiKey
  };
  var records = [];
  var page = 0;
  while (true) {
    if (options.maxPages !== null && page >= options.maxPages) {
      break;
    }
    var response = fetchJson(RECORDS_ENDPOINT, params);
    if (!response || !response.results) {
      break;
    }
    var pageResults = response.results;
    records = records.concat(pageResults);
    page++;
    logPage(options.isin, pageResults.length, page, params.offset);
    var totalCount = response.total_count || response.totalHits || 0;
    var hasMore = pageResults.length === params.limit && (params.offset + params.limit < totalCount);
    if (totalCount > 1000 && !options.dryRun) {
      var csvRecords = fetchRecordsViaCsv(selectFields, whereClause, options.apiKey, options.maxPages);
      if (csvRecords.length > records.length) {
        records = csvRecords;
      }
      break;
    }
    if (!hasMore) {
      break;
    }
    params.offset += params.limit;
    Utilities.sleep(INTER_PAGE_DELAY_MS);
  }
  var seen = {};
  var uniqueRecords = [];
  records.forEach(function (record) {
    if (!record || record.id === undefined || record.id === null) {
      return;
    }
    var recordId = String(record.id);
    if (seen[recordId]) {
      return;
    }
    seen[recordId] = true;
    uniqueRecords.push(record);
  });
  return uniqueRecords;
}

function fetchRecordsViaCsv(selectFields, whereClause, apiKey, maxPages) {
  var csvParams = {
    select: selectFields.join(','),
    where: whereClause,
    order_by: '-uin_dat_amf',
    apikey: apiKey
  };
  var response = fetchText(CSV_ENDPOINT, csvParams);
  if (!response) {
    return [];
  }
  var csv = Utilities.parseCsv(response);
  if (!csv || csv.length === 0) {
    return [];
  }
  var header = csv[0];
  var results = [];
  for (var i = 1; i < csv.length; i++) {
    if (maxPages !== null && i > 100 * maxPages) {
      break;
    }
    var row = csv[i];
    if (!row || row.length === 0) {
      continue;
    }
    var record = {};
    for (var c = 0; c < header.length; c++) {
      record[header[c]] = row[c];
    }
    record.id = record.id || record.internal_id || record.recordid || record.uuid;
    if (record.id) {
      results.push(record);
    }
  }
  return results;
}

function buildWhereClause(isin) {
  var yearStart = YEAR + '-01-01';
  var nextYearStart = (YEAR + 1) + '-01-01';
  var base = [
    "uin_dat_amf >= date'" + yearStart + "'",
    "uin_dat_amf < date'" + nextYearStart + "'",
    'emetteur_isin="' + isin.replace(/"/g, '') + '"'
  ];
  var filters = [];
  filters.push("lower(libelles) like '%document d''enregistrement universel%'");
  filters.push("lower(libelles) like '%document d enregistrement universel%'");
  filters.push("upper(libelles) like '%URD%'");
  filters.push("(lower(libelles) like '%rapport financier annuel%' AND lower(libelles) not like '%amendement%')");
  var dueFilter = '(' + filters.join(' OR ') + ')';
  base.push(dueFilter);
  return base.join(' AND ');
}

function buildSelectFields(fields) {
  var base = ['id', 'uin_dat_amf', 'emetteur_isin', 'emetteur_denomination', 'libelles'];
  if (!fields || fields.length === 0) {
    return base;
  }
  var downloadFields = [];
  fields.forEach(function (field) {
    if (DOWNLOAD_FIELD_REGEX.test(field) && base.indexOf(field) === -1) {
      downloadFields.push(field);
    }
  });
  var unique = base.concat(downloadFields.filter(function (field, index) {
    return downloadFields.indexOf(field) === index;
  }));
  return unique;
}

function fetchJson(url, params) {
  var text = fetchText(url, params, true);
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text);
  } catch (err) {
    Logger.log('Failed to parse JSON from ' + url + ': ' + err);
    return null;
  }
}

function fetchText(url, params) {
  var query = buildQuery(params);
  var fullUrl = url + (query ? '?' + query : '');
  rateLimiter.wait();
  var attempt = 0;
  while (attempt < MAX_RETRIES) {
    try {
      var response = UrlFetchApp.fetch(fullUrl, {
        muteHttpExceptions: true
      });
      var code = response.getResponseCode();
      var body = response.getContentText();
      if (code >= 200 && code < 300) {
        return body;
      }
      if ((code === 429 || code >= 500) && attempt < MAX_RETRIES - 1) {
        var backoff = Math.pow(2, attempt) * 500 + Math.floor(Math.random() * 250);
        Utilities.sleep(backoff);
        attempt++;
        continue;
      }
      Logger.log('Fetch failed ' + code + ' for ' + fullUrl + ': ' + body);
      return null;
    } catch (err) {
      if (attempt >= MAX_RETRIES - 1) {
        Logger.log('Fetch error for ' + fullUrl + ': ' + err);
        return null;
      }
      var wait = Math.pow(2, attempt) * 500 + Math.floor(Math.random() * 250);
      Utilities.sleep(wait);
      attempt++;
    }
  }
  return null;
}

function buildQuery(params) {
  if (!params) {
    return '';
  }
  var parts = [];
  Object.keys(params).forEach(function (key) {
    if (params[key] === undefined || params[key] === null || params[key] === '') {
      return;
    }
    parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(params[key])));
  });
  return parts.join('&');
}

function logPage(isin, count, page, offset) {
  Logger.log(JSON.stringify({
    time: Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX"),
    isin: isin,
    page: page,
    offset: offset,
    count: count
  }));
}

function logIssuer(isin, count, savedFiles, ms) {
  Logger.log(JSON.stringify({
    time: Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX"),
    isin: isin,
    count: count,
    savedFiles: savedFiles,
    ms: ms
  }));
}

function pickBestDownload(record) {
  var candidates = [];
  Object.keys(record).forEach(function (key) {
    var value = record[key];
    if (typeof value === 'string' && URL_REGEX.test(value)) {
      candidates.push({ field: key, url: value });
    }
  });
  if (candidates.length === 0) {
    return null;
  }
  var prioritized = candidates
    .map(function (candidate) {
      var contentType = findContentTypeForField(record, candidate.field);
      var extension = determineExtension(candidate.url, contentType);
      var priority = extensionPriority(extension);
      return {
        url: candidate.url,
        contentType: contentType,
        extension: extension,
        priority: priority
      };
    })
    .sort(function (a, b) {
      return a.priority - b.priority;
    });
  var chosen = prioritized[0];
  return {
    url: chosen.url,
    contentType: chosen.contentType || contentTypeFromExtension(chosen.extension)
  };
}

function findContentTypeForField(record, field) {
  var keys = [
    field + '_content_type',
    field + '_mime',
    field + '_format',
    field + '_type'
  ];
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (record.hasOwnProperty(key) && record[key]) {
      return String(record[key]);
    }
  }
  var altKeys = Object.keys(record).filter(function (k) {
    return k.indexOf(field) !== -1 && /mime|content_type|format/i.test(k);
  });
  for (var j = 0; j < altKeys.length; j++) {
    var altKey = altKeys[j];
    if (record[altKey]) {
      return String(record[altKey]);
    }
  }
  return null;
}

function determineExtension(url, contentType) {
  if (url) {
    var match = url.match(/\.([a-z0-9]+)(?:\?|#|$)/i);
    if (match) {
      return match[1].toLowerCase();
    }
  }
  if (contentType) {
    var lower = contentType.toLowerCase();
    if (lower.indexOf('zip') !== -1) {
      return 'zip';
    }
    if (lower.indexOf('pdf') !== -1) {
      return 'pdf';
    }
    if (lower.indexOf('html') !== -1 || lower.indexOf('xml') !== -1) {
      return 'xhtml';
    }
  }
  return 'bin';
}

function extensionPriority(extension) {
  var ext = (extension || '').toLowerCase();
  if (ext === 'zip') {
    return 0;
  }
  if (ext === 'xhtml' || ext === 'html' || ext === 'xml') {
    return 1;
  }
  if (ext === 'pdf') {
    return 2;
  }
  return 3;
}

function contentTypeFromExtension(extension) {
  var ext = (extension || '').toLowerCase();
  switch (ext) {
    case 'zip':
      return 'application/zip';
    case 'pdf':
      return 'application/pdf';
    case 'xhtml':
    case 'html':
      return 'application/xhtml+xml';
    case 'xml':
      return 'application/xml';
    default:
      return null;
  }
}

function fetchFile(url, apiKey, extension) {
  var fullUrl = url;
  if (apiKey) {
    fullUrl += (url.indexOf('?') === -1 ? '?' : '&') + 'apikey=' + encodeURIComponent(apiKey);
  }
  rateLimiter.wait();
  var attempt = 0;
  while (attempt < MAX_RETRIES) {
    try {
      var response = UrlFetchApp.fetch(fullUrl, {
        muteHttpExceptions: true
      });
      var code = response.getResponseCode();
      if (code >= 200 && code < 300) {
        var blob = response.getBlob();
        if (!blob.getContentType() && extension) {
          var type = contentTypeFromExtension(extension);
          if (type) {
            blob.setContentType(type);
          }
        }
        return blob;
      }
      if ((code === 429 || code >= 500) && attempt < MAX_RETRIES - 1) {
        var backoff = Math.pow(2, attempt) * 500 + Math.floor(Math.random() * 250);
        Utilities.sleep(backoff);
        attempt++;
        continue;
      }
      throw new Error('Download failed ' + code + ' ' + response.getContentText());
    } catch (err) {
      if (attempt >= MAX_RETRIES - 1) {
        throw err;
      }
      var wait = Math.pow(2, attempt) * 500 + Math.floor(Math.random() * 250);
      Utilities.sleep(wait);
      attempt++;
    }
  }
  throw new Error('Download failed after retries');
}

function ensureFolderPath(rootFolderId, segments) {
  var folder = DriveApp.getFolderById(rootFolderId);
  segments.forEach(function (segment) {
    var iterator = folder.getFoldersByName(segment);
    folder = iterator.hasNext() ? iterator.next() : folder.createFolder(segment);
  });
  return folder;
}

function handleZipExtraction(file, parentFolder) {
  var reportFolderIterator = parentFolder.getFoldersByName('REPORT');
  var reportFolder = reportFolderIterator.hasNext() ? reportFolderIterator.next() : parentFolder.createFolder('REPORT');
  var blobs;
  try {
    blobs = Utilities.unzip(file.getBlob());
  } catch (err) {
    Logger.log('Failed to unzip ' + file.getName() + ': ' + err);
    return;
  }
  blobs.forEach(function (blob) {
    var existing = findFileByName(reportFolder, blob.getName());
    if (existing) {
      existing.setTrashed(true);
    }
    reportFolder.createFile(blob);
  });
}

function findFileByName(folder, name) {
  var files = folder.getFilesByName(name);
  return files.hasNext() ? files.next() : null;
}

function writeMetadataFile(folder, baseName, metaJson) {
  var name = baseName + '.meta.json';
  var existing = findFileByName(folder, name);
  if (existing) {
    existing.setTrashed(true);
  }
  folder.createFile(name, metaJson, MimeType.JSON);
}

function formatDateForRecord(value) {
  if (!value) {
    return Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd');
  }
  var date = new Date(value);
  if (isNaN(date.getTime())) {
    return Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd');
  }
  return Utilities.formatDate(date, TIMEZONE, 'yyyyMMdd');
}

function makeSlug(value) {
  var text = (value || '').toLowerCase();
  text = text.replace(/[^a-z0-9]+/g, '-');
  text = text.replace(/^-+|-+$/g, '');
  return text || 'issuer';
}

function buildAmendmentSuffix(libelles) {
  var text = normalizeLibelles(libelles).toLowerCase();
  if (text.indexOf('amendement') === -1) {
    return '';
  }
  var match = normalizeLibelles(libelles).match(/amendement\s*(\d+)/i);
  var number = match ? match[1] : '1';
  return '-Amd' + number;
}

function toHex(bytes) {
  return bytes.map(function (byte) {
    var value = (byte + 256) % 256;
    var hex = value.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

function normalizeLibelles(libelles) {
  if (libelles === null || libelles === undefined) {
    return '';
  }
  if (Array.isArray(libelles)) {
    return libelles.join(' ');
  }
  if (typeof libelles === 'object') {
    try {
      return JSON.stringify(libelles);
    } catch (err) {
      return String(libelles);
    }
  }
  return String(libelles);
}
