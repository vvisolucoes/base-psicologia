function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    var data;
    try {
      data = JSON.parse(e.parameter.payload);
    } catch (ex) {
      data = JSON.parse(e.postData.contents);
    }

    var headers = [
      'data_envio', 'status', 'nome', 'idade', 'sexo', 'estado_civil', 'escolaridade',
      'ocupacao', 'reside_com', 'telefone', 'email',
      'pai_nome', 'pai_idade', 'pai_ocup', 'mae_nome', 'mae_idade', 'mae_ocup',
      'conjuge', 'conjuge_ocup', 'filhos',
      'queixa', 'inicio_sintomas', 'estressores', 'estressores_desc', 'sintomas', 'intensidade',
      'infancia', 'rel_familiar', 'obs_historia',
      'trat_anterior', 'trat_desc', 'doenca', 'doenca_desc',
      'medicamento', 'medicamento_desc', 'hist_mental', 'hist_alcool', 'hist_suicidio', 'obs_medico',
      'risco', 'observacoes_clinicas', 'sessoes', 'financeiro'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1a7a6e');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
    }

    var row = [];
    for (var i = 0; i < headers.length; i++) {
      var val = data[headers[i]];
      row.push(val !== undefined && val !== null ? val : '');
    }
    sheet.appendRow(row);

    if (data.risco && data.risco.toString().indexOf('SIM') !== -1) {
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 1, 1, headers.length).setBackground('#fde8e8');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var rows = sheet.getDataRange().getValues();

    if (rows.length < 2) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var headers = rows[0];
    var data = [];

    for (var i = 1; i < rows.length; i++) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = rows[i][j];
      }
      data.push(obj);
    }

    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
