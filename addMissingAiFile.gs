var CONFIG = {
  sheetName: 'ScrapeData',
  urlColumn: 'URL',
  aiFileIdColumn: 'AI File ID',
  aiFolderId: PropertiesService.getScriptProperties().getProperty('AI_FOLDER_ID'),
  apiEndpoint: 'https://api.openai.com/v1/chat/completions',
  apiKey: PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY'),
  model: 'gpt-3.5-turbo',
  temperature: 0.7,
  maxTokens: 1000,
  apiMaxRetries: 3,
  apiBaseDelayMs: 1000,
  perRowSleepMs: 1000
};

function addMissingAiFileFromSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error('Sheet "' + CONFIG.sheetName + '" not found.');
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var headers = data[0];
  var urlIdx = headers.indexOf(CONFIG.urlColumn);
  var fileIdIdx = headers.indexOf(CONFIG.aiFileIdColumn);
  if (urlIdx < 0 || fileIdIdx < 0) {
    throw new Error('Missing required columns: ' + CONFIG.urlColumn + ' or ' + CONFIG.aiFileIdColumn);
  }
  var folder = DriveApp.getFolderById(CONFIG.aiFolderId);
  var numRows = data.length - 1;
  var fileIds2D = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var url = row[urlIdx];
    var existingId = row[fileIdIdx];
    var newId = existingId || '';
    if (url && !existingId) {
      try {
        var summary = generateAiSummary(url);
        var timestamp = new Date().getTime();
        var fileName = slugify(url) + '-' + timestamp + '_' + (i + 1) + '.txt';
        var file = folder.createFile(fileName, summary, MimeType.PLAIN_TEXT);
        newId = file.getId();
        Utilities.sleep(CONFIG.perRowSleepMs);
      } catch (err) {
        Logger.log('Error processing row ' + (i + 1) + ': ' + err);
      }
    }
    fileIds2D.push([newId]);
  }
  var writeRange = sheet.getRange(2, fileIdIdx + 1, numRows, 1);
  writeRange.setValues(fileIds2D);
}

function generateAiSummary(url) {
  var payload = {
    model: CONFIG.model,
    messages: [
      {
        role: 'user',
        content: 'Provide a concise SEO analysis of the webpage at the following URL: ' + url
      }
    ],
    temperature: CONFIG.temperature,
    max_tokens: CONFIG.maxTokens
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  var attempt = 0;
  while (attempt < CONFIG.apiMaxRetries) {
    try {
      var response = UrlFetchApp.fetch(CONFIG.apiEndpoint, options);
      var code = response.getResponseCode();
      if (code === 200) {
        var json = JSON.parse(response.getContentText());
        var message = json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content;
        return message ? message.trim() : '';
      }
      if (code === 429 || code >= 500) {
        throw new Error('HTTP ' + code + ': ' + response.getContentText());
      }
      throw new Error('OpenAI API error ' + code + ': ' + response.getContentText());
    } catch (err) {
      attempt++;
      if (attempt >= CONFIG.apiMaxRetries) {
        throw new Error('Failed after ' + attempt + ' attempts: ' + err);
      }
      var delay = CONFIG.apiBaseDelayMs * Math.pow(2, attempt - 1);
      Utilities.sleep(delay);
    }
  }
}

function slugify(text) {
  return text
    .toString()
    .toLowerCase()
    .replace(/^https?:\/\//, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}