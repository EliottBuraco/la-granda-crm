// ============================================================================
// La Granda CRM - passerelle Google Apps Script
// Sert de base de données "API JSON" au-dessus de ce Google Sheet, pour que
// le CRM (page web publiée depuis Claude) puisse lire/écrire sans que les
// associés aient besoin d'un compte Claude ni d'un compte Google.
//
// INSTALLATION (une seule fois) :
// 1. Dans ce Google Sheet : menu "Extensions" > "Apps Script".
// 2. Supprime le contenu par défaut de "Code.gs" et colle TOUT ce fichier.
// 3. Clique sur "Déployer" (en haut à droite) > "Nouveau déploiement".
// 4. Type : "Application Web". Description : "La Granda CRM API".
// 5. "Exécuter en tant que" : Moi (ton compte Google).
// 6. "Qui a accès" : Tout le monde.
// 7. Clique "Déployer", autorise l'accès (c'est ton propre script, sur ton
//    propre Sheet), puis copie l'URL de l'application Web obtenue
//    (elle se termine par /exec).
// 8. Renvoie cette URL à Claude : c'est elle qui branche le CRM sur ce Sheet.
//
// Si tu modifies ce code plus tard, il faut aussi faire
// "Déployer" > "Gérer les déploiements" > icône crayon > "Nouvelle version"
// pour que le changement soit pris en compte par l'URL déjà donnée.
// ============================================================================

var ALLOWED_SHEETS = ["Entreprises", "Contacts", "Clients_B2C", "Objectifs", "Tournees"];

function doGet(e) {
  try {
    var action = e.parameter.action;
    var sheetName = e.parameter.collection;
    assertSheetAllowed(sheetName);

    if (action === "list") {
      return jsonOutput({ ok: true, data: listRows(sheetName) });
    }
    if (action === "get") {
      var row = getRowById(sheetName, e.parameter.id);
      return jsonOutput({ ok: true, data: row });
    }
    return jsonOutput({ ok: false, error: "unknown_action" });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sheetName = body.collection;
    assertSheetAllowed(sheetName);

    if (body.action === "create") {
      var created = createRow(sheetName, body.data || {});
      return jsonOutput({ ok: true, data: created });
    }
    if (body.action === "update") {
      var updated = updateRow(sheetName, body.id, body.data || {});
      return jsonOutput({ ok: true, data: updated });
    }
    if (body.action === "delete") {
      deleteRow(sheetName, body.id);
      return jsonOutput({ ok: true });
    }
    return jsonOutput({ ok: false, error: "unknown_action" });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function assertSheetAllowed(sheetName) {
  if (ALLOWED_SHEETS.indexOf(sheetName) === -1) {
    throw new Error("collection_not_allowed: " + sheetName);
  }
}

function getSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("sheet_not_found: " + sheetName);
  return sheet;
}

function getHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

function rowToObject(headers, rowValues) {
  var obj = {};
  for (var i = 0; i < headers.length; i++) {
    obj[headers[i]] = rowValues[i];
  }
  return obj;
}

function listRows(sheetName) {
  var sheet = getSheet(sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var headers = getHeaders(sheet);
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    var obj = rowToObject(headers, values[r]);
    if (obj.id === "" || obj.id === null) continue; // skip blank rows
    out.push(obj);
  }
  return out;
}

function findRowIndexById(sheet, headers, id) {
  var idCol = headers.indexOf("id");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // 1-indexed sheet row
  }
  return -1;
}

function getRowById(sheetName, id) {
  var sheet = getSheet(sheetName);
  var headers = getHeaders(sheet);
  var rowIndex = findRowIndexById(sheet, headers, id);
  if (rowIndex === -1) return null;
  var values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  return rowToObject(headers, values);
}

function createRow(sheetName, data) {
  var sheet = getSheet(sheetName);
  var headers = getHeaders(sheet);
  var id = Utilities.getUuid();
  var row = headers.map(function (h) {
    if (h === "id") return id;
    var v = data[h];
    return v === undefined || v === null ? "" : v;
  });
  sheet.appendRow(row);
  return rowToObject(headers, row);
}

function updateRow(sheetName, id, data) {
  var sheet = getSheet(sheetName);
  var headers = getHeaders(sheet);
  var rowIndex = findRowIndexById(sheet, headers, id);
  if (rowIndex === -1) throw new Error("not_found: " + id);
  var current = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  var next = headers.map(function (h, i) {
    if (h === "id") return current[i];
    if (Object.prototype.hasOwnProperty.call(data, h)) {
      var v = data[h];
      return v === undefined || v === null ? "" : v;
    }
    return current[i];
  });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([next]);
  return rowToObject(headers, next);
}

function deleteRow(sheetName, id) {
  var sheet = getSheet(sheetName);
  var headers = getHeaders(sheet);
  var rowIndex = findRowIndexById(sheet, headers, id);
  if (rowIndex === -1) return;
  sheet.deleteRow(rowIndex);
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
