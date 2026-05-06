/**
 * export-slides.gs
 *
 * Exports every slide in the active Google Slides presentation as a
 * high-resolution PNG and saves them into a new folder in Google Drive.
 *
 * HOW TO USE
 * ----------
 * 1. Open your Google Slides presentation.
 * 2. Extensions → Apps Script → paste this file → Save.
 * 3. Enable the Slides API as an Advanced Service:
 *      a) In the Apps Script editor, click "Services" (+ icon, left sidebar).
 *      b) Scroll to "Google Slides API" → click Add.
 * 4. Refresh google slides and wait for the "Export Slides" button and then run it. 
 * 6. Authorize when the permissions.A link to the Drive folder is shown.
 * 7. Download the folder contents for your slides
 *
 * RESOLUTION
 * ----------
 * The Slides REST API LARGE thumbnail is ~1600 × 900 px for a 16:9 deck —
 * a significant upgrade over a typical PowerPoint PNG export.
 * Each file is named Slide1.PNG, Slide2.PNG … to match the existing deck.
// ── Config ────────────────────────────────────────────────────────────────────

var FOLDER_NAME    = 'exported-slides';  // Drive folder that will be created
var FILE_PREFIX    = 'Slide';            // Slide1.PNG, Slide2.PNG …
var THUMBNAIL_SIZE = 'LARGE';           // SMALL (~200px) | MEDIUM (~400px) | LARGE (~1600px)

// ── Menu (runs automatically on open) ────────────────────────────────────────

function onOpen() {
  SlidesApp.getUi()
    .createMenu('🖼 Export Slides')
    .addItem('Export slides as PNG', 'exportSlidesToDrive')
    .addToUi();
}

// ── Main ──────────────────────────────────────────────────────────────────────

function exportSlidesToDrive() {
  var presentation  = SlidesApp.getActivePresentation();
  var presentationId = presentation.getId();
  var slides        = presentation.getSlides();
  var token         = ScriptApp.getOAuthToken();
  var folder        = DriveApp.createFolder(FOLDER_NAME + ' – ' + new Date().toISOString().slice(0, 10));

  Logger.log('Exporting %s slides to folder: %s', slides.length, folder.getName());

  var saved  = 0;
  var errors = [];

  for (var i = 0; i < slides.length; i++) {
    var slideNumber = i + 1;
    var pageId      = slides[i].getObjectId();

    // Fetch the thumbnail metadata from the Slides REST API
    var apiUrl = 'https://slides.googleapis.com/v1/presentations/'
      + presentationId
      + '/pages/'
      + pageId
      + '/thumbnail?thumbnailProperties.thumbnailSize='
      + THUMBNAIL_SIZE;

    var metaResponse = UrlFetchApp.fetch(apiUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (metaResponse.getResponseCode() !== 200) {
      var msg = 'Slide ' + slideNumber + ': HTTP ' + metaResponse.getResponseCode()
        + ' – ' + metaResponse.getContentText().slice(0, 120);
      Logger.log('ERROR ' + msg);
      errors.push(msg);
      continue;
    }

    var contentUrl = JSON.parse(metaResponse.getContentText()).contentUrl;

    // Fetch the actual PNG bytes
    var imgResponse = UrlFetchApp.fetch(contentUrl, { muteHttpExceptions: true });

    if (imgResponse.getResponseCode() !== 200) {
      var imgMsg = 'Slide ' + slideNumber + ': image fetch HTTP ' + imgResponse.getResponseCode();
      Logger.log('ERROR ' + imgMsg);
      errors.push(imgMsg);
      continue;
    }

    var blob = imgResponse.getBlob()
      .setName(FILE_PREFIX + slideNumber + '.PNG')
      .setContentType('image/png');

    folder.createFile(blob);
    saved++;
    Logger.log('Saved Slide%s.PNG (%s/%s)', slideNumber, saved, slides.length);
  }

  Logger.log('Done! %s/%s saved. Drive folder: %s', saved, slides.length, folder.getUrl());
  showCompletionAlert(folder.getUrl(), saved, slides.length, errors);
}

// ── UI helper (shows a dialog when run from the editor) ───────────────────────

function showCompletionAlert(folderUrl, saved, total, errors) {
  try {
    var ui = SlidesApp.getUi();
    var body = saved + ' of ' + total + ' slides saved.\n\nDrive folder:\n' + folderUrl;
    if (errors.length > 0) {
      body += '\n\nErrors (' + errors.length + '):\n' + errors.join('\n');
    }
    ui.alert(saved === total ? 'Export complete' : 'Export finished with errors', body, ui.ButtonSet.OK);
  } catch (e) {
    // Running headlessly – results are in View → Logs.
  }
}
