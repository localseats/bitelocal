var SHEET_ID = '1LRd2OqYOI9FHLhnShYd8pRNPQgxCuhbwaSlcvULZbnE';

// Column index (0-based): name=0, category=1, lat=2, lng=3, comment=4,
// contributor=5, timestamp=6, opening_hours=7, photos=8, instagram=9, likes=10, city=11
// Spreadsheet cols: A=name B=category C=lat D=lng E=comment F=contributor
//                   G=timestamp H=opening_hours I=photos J=instagram K=likes L=city

function doGet(e) {
  var params = e.parameter;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Spots');

  // Like action: ?action=like&row=N
  if (params && params.action === 'like' && params.row) {
    var row = parseInt(params.row);
    var cell = sheet.getRange(row, 11); // col K = likes (1-indexed)
    var current = parseInt(cell.getValue()) || 0;
    cell.setValue(current + 1);
    return ContentService.createTextOutput(JSON.stringify({ likes: current + 1 }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Write mode (GET-based fallback)
  if (params && params.name) {
    sheet.appendRow([
      params.name,
      params.category    || '',
      params.lat         || '',
      params.lng         || '',
      params.comment     || '',
      params.contributor || '',
      new Date().toISOString(),
      params.opening_hours || '',
      params.photos      || '',
      params.instagram   || '',
      0,
      params.city        || ''
    ]);
    return ContentService.createTextOutput('OK');
  }

  // Read mode
  var data = sheet.getDataRange().getValues();
  var spots = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    spots.push({
      _row:          i + 1,
      name:          data[i][0],
      category:      data[i][1],
      lat:           data[i][2],
      lng:           data[i][3],
      comment:       data[i][4],
      contributor:   data[i][5],
      timestamp:     data[i][6],
      opening_hours: data[i][7],
      photos:        data[i][8],
      instagram:     data[i][9],
      likes:         parseInt(data[i][10]) || 0,
      city:          data[i][11] || ''
    });
  }
  return ContentService.createTextOutput(JSON.stringify(spots))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // 画像アップロードモード
    if (data.action === 'uploadImage') {
      var response = UrlFetchApp.fetch('https://api.imgbb.com/1/upload', {
        method: 'post',
        payload: {
          key: '84ed403c7cd3480fc4d6753c2426e2ff',
          image: data.image
        }
      });
      var json = JSON.parse(response.getContentText());
      return ContentService.createTextOutput(JSON.stringify(json))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 通常の投稿保存モード
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Spots');
    sheet.appendRow([
      data.name          || '',
      data.category      || '',
      data.lat           || '',
      data.lng           || '',
      data.comment       || '',
      data.contributor   || '',
      new Date().toISOString(),
      data.opening_hours || '',
      data.photos        || '',
      data.instagram     || '',
      0,
      data.city          || ''
    ]);
    return ContentService.createTextOutput('OK');
  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err.toString());
  }
}
