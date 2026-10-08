/************************************************************
 * KGV NEUENHOF – BELEGVERWALTUNG
 *
 * Direkt testbare Version – Button-/Script-Fix
 * Vollständiger Code für:
 * - Formularübermittlung
 * - Beleg-ID
 * - Datei umbenennen
 * - 1. Kassierer
 * - 1. Vorsitzender
 * - 2. Kassierer
 * - Zahlungsdaten
 * - Zahlungsempfänger-Stammdaten
 * - Autocomplete
 * - Überweisung
 * - Bezahlt / Gebucht
 * - manuelle Datumsänderung
 * - Beleg löschen
 * - Belegübersicht
 * - Vorsitzendenübersicht
 * - Übersicht 2. Kassierer
 ************************************************************/


/************************************************************
 * KONFIGURATION
 ************************************************************/

const CONFIG = {

  SHEET_NAME: "Formularantworten 3",

  PAYEE_SHEET_NAME: "Zahlungsempfänger",

  // Kopien der Belege liegen im Ordner der Tabelle:
  // <Ordner der Tabelle>/Belegeingang[/Bezahlt|/Gebucht]
  BELEG_ORDNER: "Belegeingang",
  BELEG_ORDNER_BEZAHLT: "Bezahlt",
  BELEG_ORDNER_GEBUCHT: "Gebucht",
  BELEG_ORDNER_GELOESCHT: "Gelöscht",
  // Originaldateien, die über "Neuer Vorgang" hochgeladen werden
  BELEG_ORDNER_ORIGINALE: "Originale Eingabeseite",

  WEBAPP_URL:
    "https://script.google.com/a/macros/neuenhof-koeln.de/s/AKfycbyrWdMOXVOFDA0Z9hA3OCq9gZZzbc1tW1jfQeKaEIp1GNwAc3vhsllXYuuELDy5dFAk/exec",

  // TESTADRESSEN
  // Später einfach hier ändern.

  KASSIERER_EMAIL:
    "thobias.fritz@neuenhof-koeln.de",

  VORSITZENDER_EMAIL:
    "thobias.fritz@neuenhof-koeln.de",

  ZWEITER_KASSIERER_EMAIL:
    "thobias.fritz@neuenhof-koeln.de"

};


/************************************************************
 * SPALTEN
 *
 * Google Sheets zählt ab 1.
 ************************************************************/

const COL = {

  ZEITSTEMPEL: 1,
  EINREICHER: 2,
  GEKAUFT: 3,
  ZWECK: 4,
  LIEFERANT: 5,
  BETRAG: 6,
  AUSZAHLUNG_AN: 7,
  UPLOAD: 8,
  RECHNUNGSNUMMER: 9,
  EMAIL: 10,
  FREI: 11,

  BELEG_ID: 12,
  STATUS: 13,
  DATEINAME: 14,
  ORIGINALBELEG: 15,
  DATUM_ORIGINALBELEG: 16,

  ENTSCHEIDUNG_KASSIERER: 17,
  DATUM_ENTSCHEIDUNG: 18,

  GENEHMIGUNG_VORSITZENDER: 19,
  DATUM_GENEHMIGUNG: 20,
  KOMMENTAR_VORSITZENDER: 21,

  ZAHLUNGSEMPFAENGER: 22,
  IBAN: 23,
  BIC: 24,
  ZAHLUNGSZWECK: 25,
  ZAHLUNGSDATEN_GESPEICHERT: 26,

  UEBERWIESEN_AM: 27,
  INFO_UEBERWEISUNG: 28,
  KOMMENTAR_2_KASSIERER: 29,
  AN_2_KASSIERER_GESENDET: 30,

  LETZTE_BEARBEITUNG: 31,

  BEZAHLT_AM: 32,
  GEBUCHT_AM: 33,
  INTERNE_ID: 34,
  BELEGKOPIE_ID: 35,

  VORGANGSART: 36,
  ZUSATZANGABEN: 37,
  EINNAHME: 38

};

// Höchste vom Skript genutzte Spalte
const COL_MAX = 38;

// Neue Spalten bekommen beim ersten Zugriff eine Überschrift.
const COL_HEADERS = {
  34: "Interne Beleg-ID",
  35: "Belegkopie (Drive-IDs)",
  36: "Vorgangsart",
  37: "Zusatzangaben",
  38: "Einnahme"
};

const VORGANGSARTEN = [
  "Rechnung",
  "Eigenbeleg",
  "Auszahlung Kaution",
  "Auszahlung Kaution nach Nachkontrolle",
  "Quittung Wertgutachten",
  "Anderer Vorgang"
];

const VORGANGSART_STANDARD = "Rechnung";

const STATUS_EINNAHME_OFFEN = "Offen – Zahlung erwartet";
const STATUS_EINNAHME_EINGEGANGEN = "Zahlung eingegangen";



/************************************************************
 * DIAGNOSE
 ************************************************************/

function diagnoseConnection(identifier) {
  const wanted = String(identifier || "").trim();

  if (!wanted) {
    return {
      ok: false,
      message: "Kein Beleg-Schlüssel übergeben.",
      identifier: ""
    };
  }

  const row = findRowByIdentifier(wanted);

  if (!row) {
    return {
      ok: false,
      message: "Server erreichbar, aber Beleg nicht gefunden.",
      identifier: wanted
    };
  }

  const sheet = getSheet();
  const actualBelegId = String(
    sheet.getRange(row, COL.BELEG_ID).getValue() || ""
  ).trim();

  return {
    ok: true,
    message: "Verbindung zum Apps-Script-Server funktioniert.",
    identifier: wanted,
    row: row,
    belegId: actualBelegId
  };
}


/************************************************************
 * HILFSFUNKTIONEN
 ************************************************************/


function getSheet() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error(
      "Das Tabellenblatt '" +
      CONFIG.SHEET_NAME +
      "' wurde nicht gefunden."
    );
  }

  ensureColumns(sheet);

  return sheet;
}


// Stellt sicher, dass alle vom Skript genutzten Spalten existieren
// und eine Überschrift haben (einmal pro Skriptausführung).
let columnsChecked = false;

function ensureColumns(sheet) {

  if (columnsChecked) {
    return;
  }

  if (sheet.getMaxColumns() < COL_MAX) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      COL_MAX - sheet.getMaxColumns()
    );
  }

  const headerRange = sheet.getRange(1, 1, 1, COL_MAX);
  const headers = headerRange.getValues()[0];
  let changed = false;

  Object.keys(COL_HEADERS).forEach(function(col) {
    if (!String(headers[col - 1] || "").trim()) {
      headers[col - 1] = COL_HEADERS[col];
      changed = true;
    }
  });

  if (changed) {
    headerRange.setValues([headers]);
  }

  columnsChecked = true;

}


// Überschrift einer Spalte aus Zeile 1 (z. B. für Spalte I,
// deren Bedeutung im Formular festgelegt wird).
function getColumnHeader(col, fallback) {

  const value =
    String(getSheet().getRange(1, col).getValue() || "").trim();

  return value || fallback || "";

}


// Betrag aus Formular/Eingabe in eine Zahl umwandeln.
// Erlaubt: 12,50 · 12.50 · 1.234,56 · 12,50 €
// Leere Eingabe ergibt "". Ungültige Eingabe wirft einen Fehler.
function parseAmount(value) {

  if (typeof value === "number") {
    return value;
  }

  let text =
    String(value || "")
      .replace(/[€\s]/g, "")
      .trim();

  if (!text) {
    return "";
  }

  if (text.indexOf(",") >= 0) {
    text = text.replace(/\./g, "").replace(",", ".");
  }

  const number = Number(text);

  if (isNaN(number) || number < 0) {
    throw new Error("Ungültiger Betrag: " + value);
  }

  return Math.round(number * 100) / 100;

}


function isEinnahme(values) {
  return String(values[COL.EINNAHME - 1] || "").trim() === "Ja";
}


function getVorgangsart(values) {
  return String(values[COL.VORGANGSART - 1] || "").trim() || VORGANGSART_STANDARD;
}


// Betrag darf nicht mehr geändert werden, sobald überwiesen wurde.
function isAmountLocked(values) {

  if (values[COL.UEBERWIESEN_AM - 1]) {
    return true;
  }

  const status = String(values[COL.STATUS - 1] || "");

  return [
    "Überwiesen",
    "Bezahlt",
    "Gebucht",
    STATUS_EINNAHME_EINGEGANGEN
  ].indexOf(status) >= 0;

}


/************************************************************
 * NÄCHSTE BELEG-ID (JJJJ-NNNN)
 *
 * Nur innerhalb eines Script-Locks aufrufen.
 * Die Nummer wird aus der höchsten vorhandenen Nummer des
 * Jahres abgeleitet, nicht aus der Zeilennummer.
 ************************************************************/

function nextBelegId(sheet, year) {

  const lastRow =
    sheet.getLastRow();

  const existingIds =
    lastRow >= 2
      ? sheet
          .getRange(2, COL.BELEG_ID, lastRow - 1, 1)
          .getValues()
      : [];

  let highestNumber = 0;

  existingIds.forEach(function(item) {

    const match =
      String(item[0] || "").trim()
        .match(new RegExp("^(\\d{4})-(\\d+)$"));

    if (!match || Number(match[1]) !== Number(year)) {
      return;
    }

    highestNumber =
      Math.max(highestNumber, Number(match[2]));

  });

  return (
    year +
    "-" +
    Utilities.formatString("%04d", highestNumber + 1)
  );

}


function getOrCreateInternalId(rowNumber) {
  const sheet = getSheet();
  let value = String(sheet.getRange(rowNumber, COL.INTERNE_ID).getValue() || "").trim();
  if (!value) { value = Utilities.getUuid(); sheet.getRange(rowNumber, COL.INTERNE_ID).setValue(value); }
  return value;
}

function ensureAllInternalIds() {
  const sheet = getSheet();
  if (!sheet.getRange(1, COL.INTERNE_ID).getValue()) sheet.getRange(1, COL.INTERNE_ID).setValue("Interne Beleg-ID");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, COL.INTERNE_ID, lastRow - 1, 1);
  const values = range.getValues();
  let changed = false;
  for (let i=0;i<values.length;i++) { if (!String(values[i][0] || "").trim()) { values[i][0]=Utilities.getUuid(); changed=true; } }
  if (changed) range.setValues(values);
}

function findRowByInternalId(internalId) {
  const sheet=getSheet(); const lastRow=sheet.getLastRow(); if(lastRow<2 || !internalId) return null;
  const ids=sheet.getRange(2,COL.INTERNE_ID,lastRow-1,1).getValues(); const wanted=String(internalId).trim();
  for(let i=0;i<ids.length;i++) if(String(ids[i][0]||"").trim()===wanted) return i+2;
  return null;
}

function findRowByBelegId(belegId) {

  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !belegId) return null;

  const wanted = String(belegId).trim();

  // Zuerst nach der normalen Beleg-ID suchen.
  const ids = sheet.getRange(2, COL.BELEG_ID, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === wanted) {
      return i + 2;
    }
  }

  // Bei ?key=... kann die uebergebene Kennung die interne UUID sein.
  // Deshalb auch die interne ID durchsuchen.
  if (COL.INTERNE_ID) {
    const internalIds = sheet.getRange(2, COL.INTERNE_ID, lastRow - 1, 1).getValues();
    for (let i = 0; i < internalIds.length; i++) {
      if (String(internalIds[i][0] || '').trim() === wanted) {
        return i + 2;
      }
    }
  }

  return null;
}

function findRowByIdentifier(identifier) {
  const wanted = String(identifier || "").trim();
  if (!wanted) return null;

  // Neue Links verwenden die interne UUID (?key=...).
  // Alte Links verwenden weiterhin die normale Beleg-ID (?id=...).
  const internalRow = findRowByInternalId(wanted);
  if (internalRow) return internalRow;

  return findRowByBelegId(wanted);
}



function getRowData(identifier, useInternalId) {
  const sheet=getSheet();
  const rowNumber=useInternalId ? findRowByInternalId(identifier) : findRowByBelegId(identifier);
  if(!rowNumber) return null;
  const values=sheet.getRange(rowNumber,1,1,COL_MAX).getValues()[0];
  getOrCreateInternalId(rowNumber);
  values[COL.INTERNE_ID-1]=sheet.getRange(rowNumber,COL.INTERNE_ID).getValue();
  return {rowNumber:rowNumber,values:values};
}


function formatDate(value) {

  if (!value) {
    return "";
  }

  try {

    return Utilities.formatDate(
      new Date(value),
      Session.getScriptTimeZone(),
      "dd.MM.yyyy"
    );

  } catch (e) {

    return "";

  }

}


function formatDateTime(value) {

  if (!value) {
    return "";
  }

  try {

    return Utilities.formatDate(
      new Date(value),
      Session.getScriptTimeZone(),
      "dd.MM.yyyy HH:mm"
    );

  } catch (e) {

    return "";

  }

}


function htmlEscape(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

}


function jsonForHtml(value) {
  return JSON.stringify(value)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
}


function makeWebUrl(belegId, extraParameters) {
  const row=findRowByIdentifier(belegId); let url=CONFIG.WEBAPP_URL;
  if(row) url += "?key="+encodeURIComponent(getOrCreateInternalId(row));
  else url += "?id="+encodeURIComponent(belegId);
  if(extraParameters) Object.keys(extraParameters).forEach(function(key){ url += "&"+encodeURIComponent(key)+"="+encodeURIComponent(extraParameters[key]); });
  return url;
}


function sendMail(
  recipient,
  subject,
  body
) {

  MailApp.sendEmail({

    to: recipient,

    subject: subject,

    body: body

  });

}


function getStatusClass(status) {

  const s =
    String(status || "")
      .toLowerCase();

  if (
    s.indexOf("neu") >= 0 ||
    s.indexOf("offen") >= 0
  ) {
    return "status-neu";
  }

  if (
    s.indexOf("genehmigung") >= 0
  ) {
    return "status-genehmigung";
  }

  if (
    s.indexOf("genehmigt") >= 0
  ) {
    return "status-genehmigt";
  }

  if (
    s.indexOf("zahlung") >= 0
  ) {
    return "status-zahlung";
  }

  if (
    s.indexOf("überwiesen") >= 0
  ) {
    return "status-ueberwiesen";
  }

  if (
    s.indexOf("bezahlt") >= 0
  ) {
    return "status-bezahlt";
  }

  if (
    s.indexOf("gebucht") >= 0
  ) {
    return "status-gebucht";
  }

  if (
    s.indexOf("abgelehnt") >= 0
  ) {
    return "status-abgelehnt";
  }

  if (
    s.indexOf("ablage") >= 0
  ) {
    return "status-ablage";
  }

  return "status-standard";
}


/************************************************************
 * FORMULARÜBERMITTLUNG
 ************************************************************/


function onFormSubmit(e) {

  const sheet =
    e.range.getSheet();

  const row =
    e.range.getRow();

  if (
    sheet.getName() !==
    CONFIG.SHEET_NAME
  ) {
    return;
  }


  const supplier =
    sheet
      .getRange(
        row,
        COL.LIEFERANT
      )
      .getValue();

  const amount =
    sheet
      .getRange(
        row,
        COL.BETRAG
      )
      .getValue();

  const purpose =
    sheet
      .getRange(
        row,
        COL.ZWECK
      )
      .getValue();

  const uploadLink =
    sheet
      .getRange(
        row,
        COL.UPLOAD
      )
      .getValue();


  /******************************************************
   * BELEG-ID
   ******************************************************/

  /******************************************************
   * BELEG-ID – EINDEUTIG UND SICHER
   *
   * Die Beleg-ID darf NICHT aus der Zeilennummer
   * abgeleitet werden. Gelöschte Zeilen könnten sonst
   * dazu führen, dass eine alte Beleg-ID erneut vergeben
   * wird.
   *
   * LockService verhindert zusätzlich, dass zwei nahezu
   * gleichzeitig eingehende Formulare dieselbe Nummer
   * bekommen.
   ******************************************************/

  const lock =
    LockService.getScriptLock();

  lock.waitLock(30000);

  let belegId;

  try {

    const timestamp =
      sheet
        .getRange(row, COL.ZEITSTEMPEL)
        .getValue();

    const year =
      timestamp instanceof Date && !isNaN(timestamp.getTime())
        ? timestamp.getFullYear()
        : new Date().getFullYear();

    belegId =
      nextBelegId(sheet, year);
    sheet
      .getRange(row, COL.BELEG_ID)
      .setValue(belegId);

    sheet
      .getRange(row, COL.STATUS)
      .setValue("Neu");

    if (!sheet.getRange(row, COL.VORGANGSART).getValue()) {
      sheet
        .getRange(row, COL.VORGANGSART)
        .setValue(VORGANGSART_STANDARD);
    }

  } finally {

    lock.releaseLock();

  }


  if (!sheet.getRange(1, COL.INTERNE_ID).getValue()) sheet.getRange(1, COL.INTERNE_ID).setValue("Interne Beleg-ID");
  const internalId = getOrCreateInternalId(row);


  /******************************************************
   * DATEI UMBENENNEN
   ******************************************************/

  try {

    if (
      uploadLink &&
      String(uploadLink).indexOf("id=") >= 0
    ) {

      const match =
        String(uploadLink)
          .match(
            /id=([^&]+)/
          );

      if (match) {

        const fileId =
          match[1];

        const file =
          DriveApp.getFileById(
            fileId
          );

        const oldName =
          file.getName();

        let extension =
          "";

        if (
          oldName.indexOf(".") >= 0
        ) {

          extension =
            "." +
            oldName
              .split(".")
              .pop();

        }

        const safeSupplier =
          String(supplier || "Beleg")
            .replace(
              /[^\wäöüÄÖÜß-]/g,
              "_"
            );

        const newFilename =
          belegId +
          "_" +
          safeSupplier +
          extension;

        file.setName(
          newFilename
        );

        sheet
          .getRange(
            row,
            COL.DATEINAME
          )
          .setValue(
            newFilename
          );

      }

    }

  } catch (error) {

    sheet
      .getRange(
        row,
        COL.DATEINAME
      )
      .setValue(
        "Dateifehler"
      );

    Logger.log(
      error
    );

  }


  /******************************************************
   * KOPIE IN "BELEGEINGANG"
   ******************************************************/

  syncBelegCopiesSafe(row);


  /******************************************************
   * E-MAIL 1. KASSIERER
   ******************************************************/

  const url =
    makeWebUrl(
      belegId
    );

  const subject =
    "Neue Rechnung zur Prüfung – " +
    belegId;

  const body =
    "Eine neue Rechnung wurde eingereicht.\n\n" +

    "Beleg-ID: " +
    belegId +
    "\n" +

    "Lieferant: " +
    supplier +
    "\n" +

    "Betrag: " +
    amount +
    " €\n" +

    "Verwendungszweck: " +
    purpose +
    "\n\n" +

    "Bitte prüfe den Beleg:\n" +
    url;

  sendMail(
    CONFIG.KASSIERER_EMAIL,
    subject,
    body
  );

}


/************************************************************
 * WEB-APP ROUTER
 ************************************************************/


function doGet(e) {

  const params =
    e &&
    e.parameter
      ? e.parameter
      : {};

  const belegId=params.id || "";
  const internalKey=params.key || "";

  const rolle =
    params.rolle || "";

  const uebersicht =
    params.uebersicht || "";


  if (params.neu === "1") {

    return HtmlService
      .createHtmlOutput(
        createNewEntryPage()
      )
      .setTitle(
        "Neuer Vorgang – KGV Neuenhof"
      );

  }


  if (
    uebersicht === "1"
  ) {

    if (
      rolle === "vorsitzender"
    ) {

      return HtmlService
        .createHtmlOutput(
          createChairmanOverviewPage()
        )
        .setTitle(
          "Genehmigungen – KGV Neuenhof"
        );

    }

    if (
      rolle === "zweiter_kassierer"
    ) {

      return HtmlService
        .createHtmlOutput(
          createPaymentOverview()
        )
        .setTitle(
          "Meine Überweisungen – KGV Neuenhof"
        );

    }

    return HtmlService
      .createHtmlOutput(
        createOverviewPage()
      )
      .setTitle(
        "Belegübersicht – KGV Neuenhof"
      );

  }


  if (!belegId && !internalKey) {

    return HtmlService
      .createHtmlOutput(
        createStartPage()
      )
      .setTitle(
        "Belegprüfung KGV Neuenhof"
      );

  }


  if (
    rolle === "vorsitzender"
  ) {

    return HtmlService
      .createHtmlOutput(
        createChairmanPage(belegId || internalKey, !!internalKey)
      )
      .setTitle(
        "Beleggenehmigung – KGV Neuenhof"
      );

  }


  if (
    rolle === "zweiter_kassierer"
  ) {

    return HtmlService
      .createHtmlOutput(
        createPaymentPage(belegId || internalKey, !!internalKey)
      )
      .setTitle(
        "Überweisung – KGV Neuenhof"
      );

  }


  return HtmlService
    .createHtmlOutput(
      createDetailPage(
        belegId || internalKey,
        !!internalKey
      )
    )
    .setTitle(
      "Belegprüfung KGV Neuenhof"
    );

}


/************************************************************
 * STARTSEITE
 ************************************************************/


function createStartPage() {

  return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<style>

body {
  font-family: Arial, sans-serif;
  background: #f5f7fa;
  padding: 20px;
}

.container {
  max-width: 700px;
  margin: auto;
  background: white;
  padding: 30px;
  border-radius: 12px;
  box-shadow: 0 2px 10px rgba(0,0,0,.12);
}

.button {
  display: block;
  width: 100%;
  box-sizing: border-box;
  padding: 14px;
  margin-top: 15px;
  border-radius: 8px;
  border: none;
  text-align: center;
  text-decoration: none;
  font-weight: bold;
  font-size: 16px;
  cursor: pointer;
  background: #1769aa;
  color: white;
}

</style>

</head>

<body>

<div class="container">

<h1>Belegprüfung KGV Neuenhof</h1>

<p>
Bitte öffne einen Beleg über den persönlichen Link
aus der E-Mail.
</p>

<a
  class="button"
  href="${CONFIG.WEBAPP_URL}?uebersicht=1"
  target="_blank"
>
Belegübersicht öffnen
</a>

</div>

</body>

</html>

`;

}


/************************************************************
 * GEMEINSAMES CSS
 ************************************************************/


function commonCss() {

  return `

body {
  font-family: Arial, sans-serif;
  background: #f5f7fa;
  margin: 0;
  padding: 20px;
}

.container {
  max-width: 850px;
  margin: auto;
  background: white;
  padding: 25px;
  border-radius: 12px;
  box-shadow: 0 2px 10px rgba(0,0,0,.12);
}

h1 {
  margin-top: 0;
}

.beleg-id {
  font-size: 24px;
  font-weight: bold;
  margin-bottom: 15px;
}

table {
  width: 100%;
  border-collapse: collapse;
}

td, th {
  padding: 10px;
  border-bottom: 1px solid #ddd;
  text-align: left;
  vertical-align: top;
}

.label {
  font-weight: bold;
  width: 35%;
}

.button {
  display: block;
  width: 100%;
  box-sizing: border-box;
  text-align: center;
  padding: 14px;
  margin-top: 12px;
  border-radius: 8px;
  border: none;
  text-decoration: none;
  font-weight: bold;
  font-size: 16px;
  cursor: pointer;
}

.button:disabled {
  background: #999 !important;
  cursor: default;
}

.document-button {
  background: #1769aa;
  color: white;
}

.original-button {
  background: #188038;
  color: white;
}

.decision-button {
  background: #5f6368;
  color: white;
}

.approve-button {
  background: #188038;
  color: white;
}

.reject-button {
  background: #c5221f;
  color: white;
}

.delete-button {
  background: #c5221f;
  color: white;
}

.payment-button {
  background: #1769aa;
  color: white;
}

.save-button {
  background: #188038;
  color: white;
}

.secondary-button {
  background: #5f6368;
  color: white;
}

.back-button {
  background: #444;
  color: white;
}

.section {
  margin-top: 25px;
  padding: 20px;
  background: #f8f9fa;
  border-radius: 10px;
}

.section-title {
  font-size: 19px;
  font-weight: bold;
  margin-bottom: 15px;
}

.status-box {
  margin-top: 20px;
  padding: 15px;
  border-radius: 8px;
  font-weight: bold;
}

.status-neu {
  background: #fff3cd;
}

.status-genehmigung {
  background: #fff3cd;
}

.status-genehmigt {
  background: #d9ead3;
}

.status-zahlung {
  background: #d9eaf7;
}

.status-ueberwiesen {
  background: #d9eaf7;
}

.status-bezahlt {
  background: #d9ead3;
}

.status-gebucht {
  background: #b6d7a8;
}

.status-abgelehnt {
  background: #f4cccc;
}

.status-ablage {
  background: #e6e6e6;
}

.status-standard {
  background: #f1f3f4;
}

input,
textarea,
select {
  width: 100%;
  box-sizing: border-box;
  padding: 11px;
  border: 1px solid #ccc;
  border-radius: 7px;
  font-size: 16px;
  margin-top: 5px;
}

textarea {
  min-height: 100px;
  resize: vertical;
}

.form-row {
  margin-bottom: 15px;
}

.form-label {
  font-weight: bold;
}

.checkbox-row {
  margin-top: 15px;
}

.checkbox-row input {
  width: auto;
}

.message {
  margin-top: 15px;
  padding: 12px;
  border-radius: 7px;
  background: #d9ead3;
}

.hidden {
  display: none;
}

.warning {
  margin-top: 15px;
  padding: 12px;
  border-radius: 7px;
  background: #fff3cd;
}

`;

}


/************************************************************
 * 1. KASSIERER – DETAILSEITE
 ************************************************************/


function createDetailPage(
  belegId, useInternalId
) {

  const result =
    getRowData(belegId, useInternalId);

  if (!result) {

    return errorPage(
      "Der Beleg " +
      htmlEscape(belegId) +
      " wurde nicht gefunden."
    );

  }

  const row =
    result.values;


  const timestamp =
    formatDateTime(
      row[COL.ZEITSTEMPEL - 1]
    );

  const einreicher =
    row[COL.EINREICHER - 1];

  const gekauft =
    row[COL.GEKAUFT - 1];

  const zweck =
    row[COL.ZWECK - 1];

  const lieferant =
    row[COL.LIEFERANT - 1];

  const betrag =
    row[COL.BETRAG - 1];

  const auszahlungAn =
    row[COL.AUSZAHLUNG_AN - 1];

  const upload =
    row[COL.UPLOAD - 1];

  const rechnungsnummer =
    row[COL.RECHNUNGSNUMMER - 1];

  const email =
    row[COL.EMAIL - 1];

  const status =
    row[COL.STATUS - 1];

  const dateiname =
    row[COL.DATEINAME - 1];

  const originalbeleg =
    row[COL.ORIGINALBELEG - 1];

  const originalDatum =
    formatDate(
      row[COL.DATUM_ORIGINALBELEG - 1]
    );

  const entscheidung =
    row[COL.ENTSCHEIDUNG_KASSIERER - 1];

  const entscheidungsDatum =
    formatDate(
      row[COL.DATUM_ENTSCHEIDUNG - 1]
    );

  const genehmigung =
    row[COL.GENEHMIGUNG_VORSITZENDER - 1];

  const genehmigungDatum =
    formatDate(
      row[COL.DATUM_GENEHMIGUNG - 1]
    );

  const kommentarVorsitzender =
    row[COL.KOMMENTAR_VORSITZENDER - 1];

  const zahlungsempfaenger =
    row[COL.ZAHLUNGSEMPFAENGER - 1];

  const iban =
    row[COL.IBAN - 1];

  const bic =
    row[COL.BIC - 1];

  const zahlungszweck =
    row[COL.ZAHLUNGSZWECK - 1];

  const zahlungsdatenGespeichert =
    formatDate(
      row[COL.ZAHLUNGSDATEN_GESPEICHERT - 1]
    );

  const ueberwiesenAm =
    formatDate(
      row[COL.UEBERWIESEN_AM - 1]
    );

  const infoUeberweisung =
    row[COL.INFO_UEBERWEISUNG - 1];

  const kommentar2Kassierer =
    row[COL.KOMMENTAR_2_KASSIERER - 1];

  const an2Kassierer =
    formatDate(
      row[COL.AN_2_KASSIERER_GESENDET - 1]
    );

  const bezahltAm =
    formatDate(
      row[COL.BEZAHLT_AM - 1]
    );

  const gebuchtAm =
    formatDate(
      row[COL.GEBUCHT_AM - 1]
    );

  const einnahme =
    isEinnahme(row);

  const bezahltLabel =
    einnahme ? "Eingegangen am" : "Bezahlt am";


  let originalAnzeige =
    originalbeleg === "Ja"
      ? "✅ Originalbeleg vorhanden"
      : "❌ Originalbeleg noch nicht bestätigt";


  let entscheidungsAnzeige =
    entscheidung
      ? "Entscheidung: " +
        htmlEscape(entscheidung) +
        "<br>Datum: " +
        htmlEscape(entscheidungsDatum)
      : "Noch keine Entscheidung getroffen.";


  const payeeData =
    getPaymentRecipients();


  return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Belegprüfung KGV Neuenhof</title>

<style>

${commonCss()}

.payee-wrapper {
  position: relative;
}

.payee-suggestions {
  position: absolute;
  z-index: 50;
  left: 0;
  right: 0;
  background: white;
  border: 1px solid #ccc;
  border-radius: 0 0 7px 7px;
  max-height: 220px;
  overflow-y: auto;
  display: none;
}

.payee-suggestion {
  padding: 10px;
  cursor: pointer;
  border-bottom: 1px solid #eee;
}

.payee-suggestion:hover {
  background: #f1f3f4;
}

.small {
  font-size: 13px;
  color: #666;
}

</style>

<script>

const belegId =
  ${jsonForHtml(String(result.values[COL.BELEG_ID - 1] || belegId))};

const rowNumber =
  ${jsonForHtml(result.rowNumber)};

let sepaAmount =
  ${jsonForHtml(String(betrag || "").replace(/[^\d,.-]/g, "").replace(",", "."))};

const payees =
  ${jsonForHtml(payeeData)};



function buildSepaPayload() {
  const recipient = document.getElementById("zahlungsempfaenger");
  const ibanEl = document.getElementById("iban");
  const bicEl = document.getElementById("bic");
  const purposeEl = document.getElementById("zahlungszweck");

  if (!recipient || !ibanEl || !purposeEl) return "";

  const name = recipient.value.trim();
  const iban = ibanEl.value.replace(/\\s+/g, "").trim();
  const bic = bicEl ? bicEl.value.replace(/\\s+/g, "").trim() : "";
  const purpose = purposeEl.value.trim();

  if (!name || !iban || !purpose) return "";

  const value = Number(String(sepaAmount || "").replace(",", ".").trim());

  // Aufbau nach EPC-Standard (GiroCode). Der Verwendungszweck
  // gehört in Zeile 11, die Zeilen 9 und 10 bleiben leer.
  return [
    "BCD",
    "002",
    "1",
    "SCT",
    bic.toUpperCase(),
    name.substring(0, 70),
    iban.toUpperCase(),
    value > 0 ? "EUR" + value.toFixed(2) : "",
    "",
    "",
    purpose.substring(0, 140)
  ].join("\\n");
}

function onFormDataSaved(result) {
  sepaAmount = result.amountRaw;
  renderSepaQrCode();
}


function renderSepaQrCode() {
  const box = document.getElementById("sepaQrCode");
  if (!box) return;

  const payload = buildSepaPayload();

  if (!payload) {
    box.innerHTML =
      "<div class='small'>Bitte Zahlungsempfänger, IBAN und Verwendungszweck eintragen.</div>";
    return;
  }

  box.innerHTML =
    "<img alt='SEPA QR-Code' " +
    "style='max-width:320px;width:100%;height:auto' " +
    "src='https://quickchart.io/qr?text=" +
    encodeURIComponent(payload) +
    "&size=320'>" +
    "<div class='small' style='margin-top:8px'>SEPA-Überweisungsdaten</div>";
}


function showMessage(
  text
) {

  const el =
    document.getElementById(
      "actionMessage"
    );

  if (el) {

    el.innerHTML =
      "✅ " + text;

    el.style.display =
      "block";

  }

}


function originalbelegVorhanden() {

  const button =
    document.getElementById(
      "originalButton"
    );

  if (!button) {
    return;
  }

  button.disabled = true;

  button.innerText =
    "Wird gespeichert...";

  google.script.run

    .withSuccessHandler(
      function(result) {

        button.innerText =
          "✓ Originalbeleg vorhanden";

        document.getElementById(
          "originalStatus"
        ).innerHTML =
          "✅ " + result;

        document.getElementById(
          "originalConfirmed"
        ).checked = true;

      }
    )

    .withFailureHandler(
      function(error) {

        button.disabled = false;

        button.innerText =
          "✓ Originalbeleg vorhanden";

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .originalbelegVorhandenByRow(
      rowNumber
    );

}


function collectPaymentData() {

  return {

    zahlungsempfaenger:
      document.getElementById(
        "zahlungsempfaenger"
      ).value,

    iban:
      document.getElementById(
        "iban"
      ).value,

    bic:
      document.getElementById(
        "bic"
      ).value,

    zahlungszweck:
      document.getElementById(
        "zahlungszweck"
      ).value,

    info:
      document.getElementById(
        "info"
      ).value,

    masterUpdate:
      document.getElementById(
        "masterUpdate"
      ).checked

  };

}


function savePaymentData() {

  const data =
    collectPaymentData();


  const button =
    document.getElementById(
      "savePaymentButton"
    );

  button.disabled = true;

  button.innerText =
    "Wird gespeichert...";


  google.script.run

    .withSuccessHandler(
      function(result) {

        button.disabled =
          false;

        button.innerText =
          "💾 Zahlungsdaten speichern";

        showMessage(
          result
        );

      }
    )

    .withFailureHandler(
      function(error) {

        button.disabled =
          false;

        button.innerText =
          "💾 Zahlungsdaten speichern";

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .savePaymentDataByRow(
      rowNumber,
      data
    );

}


function sendToSecondCashier() {

  if (
    !document.getElementById(
      "originalConfirmed"
    ).checked
  ) {

    alert(
      "Bitte zuerst bestätigen, dass der Originalbeleg vorhanden ist."
    );

    return;

  }


  const recipient =
    document.getElementById(
      "zahlungsempfaenger"
    ).value.trim();

  const iban =
    document.getElementById(
      "iban"
    ).value.trim();

  const purpose =
    document.getElementById(
      "zahlungszweck"
    ).value.trim();


  if (!recipient) {

    alert(
      "Bitte einen Zahlungsempfänger eintragen."
    );

    return;

  }

  if (!iban) {

    alert(
      "Bitte eine IBAN eintragen."
    );

    return;

  }

  if (!purpose) {

    alert(
      "Bitte einen Zahlungsverwendungszweck eintragen."
    );

    return;

  }


  const button =
    document.getElementById(
      "sendPaymentButton"
    );

  button.disabled = true;

  button.innerText =
    "Wird übergeben...";


  google.script.run

    .withSuccessHandler(
      function(result) {

        button.innerText =
          "✓ An 2. Kassierer übergeben";

        showMessage(
          result
        );

      }
    )

    .withFailureHandler(
      function(error) {

        button.disabled =
          false;

        button.innerText =
          "→ An 2. Kassierer zur Überweisung";

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .sendToSecondCashierByRow(
      rowNumber,
      Object.assign(
        collectPaymentData(),
        {originalConfirmed: true}
      )
    );

}


function updatePayeeSuggestions() {

  const input = document.getElementById("zahlungsempfaenger");
  const box = document.getElementById("payeeSuggestions");

  if (!input || !box) return;

  const query = input.value.trim();

  box.innerHTML = "";

  if (!query) {
    box.style.display = "none";
    return;
  }

  // Die Stammdaten werden bei jeder Eingabe aktuell vom Server abgefragt.
  // Dadurch funktionieren die Vorschläge auch bei neu angelegten/aktualisierten Empfängern.
  google.script.run
    .withSuccessHandler(function(matches) {

      // Nur die Antwort für den aktuell eingegebenen Text anzeigen.
      if (input.value.trim() !== query) return;

      box.innerHTML = "";

      if (!matches || !matches.length) {
        box.style.display = "none";
        return;
      }

      matches.forEach(function(item) {

        const div = document.createElement("div");
        div.className = "payee-suggestion";

        div.innerHTML =
          "<strong>" + escapeHtml(item.name) + "</strong><br>" +
          "<span class='small'>" +
          escapeHtml(item.iban || "") +
          (item.bic ? " · " + escapeHtml(item.bic) : "") +
          "</span>";

        div.onclick = function() {

          document.getElementById("zahlungsempfaenger").value = item.name;
          document.getElementById("iban").value = item.iban || "";
          document.getElementById("bic").value = item.bic || "";

          box.style.display = "none";

          // Nach Auswahl den SEPA-QR-Code sofort aktualisieren.
          if (typeof renderSepaQrCode === "function") {
            renderSepaQrCode();
          }
        };

        box.appendChild(div);
      });

      box.style.display = "block";
    })
    .withFailureHandler(function(error) {
      box.style.display = "none";
      console.error("Fehler bei Zahlungsempfänger-Suche:", error);
    })
    .getPayeeSuggestions(query);

}


function escapeHtml(
  value
) {

  return String(value || "")
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;")
    .replace(/'/g,"&#039;");

}


// Das Skript steht im <head>. Die Eingabefelder existieren erst nach dem
// Laden der Seite, daher die Listener erst dann registrieren.
window.addEventListener("load", function() {
  ["zahlungsempfaenger","iban","bic","zahlungszweck"].forEach(function(id) {
    const field = document.getElementById(id);
    if (field) {
      field.addEventListener("input", renderSepaQrCode);
    }
  });
  renderSepaQrCode();
});

document.addEventListener(
  "click",
  function(event) {

    if (
      !event.target.closest(
        ".payee-wrapper"
      )
    ) {

      const box =
        document.getElementById(
          "payeeSuggestions"
        );

      if (box) {
        box.style.display =
          "none";
      }

    }

  }
);


function editAccountingDates() {

  const bezahlt =
    document.getElementById(
      "bezahltDatum"
    ).value;

  const gebucht =
    document.getElementById(
      "gebuchtDatum"
    ).value;


  google.script.run

    .withSuccessHandler(
      function(result) {

        showMessage(
          result
        );

      }
    )

    .withFailureHandler(
      function(error) {

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .setAccountingDatesByRow(
      rowNumber,
      {bezahlt: bezahlt, gebucht: gebucht}
    );

}


function markAsPaid() {

  let date =
    document.getElementById(
      "bezahltDatum"
    ).value;


  if (!date) {
    const today = new Date();
    date =
      today.getFullYear() + "-" +
      String(today.getMonth() + 1).padStart(2, "0") + "-" +
      String(today.getDate()).padStart(2, "0");
    document.getElementById("bezahltDatum").value = date;
  }


  google.script.run

    .withSuccessHandler(
      function(result) {

        showMessage(
          result
        );

      }
    )

    .withFailureHandler(
      function(error) {

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .markAsPaidByRow(
      rowNumber,
      date
    );

}


function markAsBooked() {

  let date =
    document.getElementById(
      "gebuchtDatum"
    ).value;


  if (!date) {
    const today = new Date();
    date =
      today.getFullYear() + "-" +
      String(today.getMonth() + 1).padStart(2, "0") + "-" +
      String(today.getDate()).padStart(2, "0");
    document.getElementById("gebuchtDatum").value = date;
  }


  google.script.run

    .withSuccessHandler(
      function(result) {

        showMessage(
          result
        );

      }
    )

    .withFailureHandler(
      function(error) {

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .markAsBookedByRow(
      rowNumber,
      date
    );

}


function deleteBeleg() {

  const answer =
    confirm(
      "Wirklich löschen?\\n\\n" +
      "Der Beleg " +
      belegId +
      " wird aus der Belegliste gelöscht."
    );


  if (!answer) {
    return;
  }


  const button =
    document.getElementById(
      "deleteButton"
    );

  button.disabled =
    true;

  button.innerText =
    "Wird gelöscht...";


  google.script.run

    .withSuccessHandler(
      function(result) {

        document.body.innerHTML =
          "<div style='font-family:Arial;padding:30px;text-align:center'>" +
          "<h1>Beleg gelöscht</h1>" +
          "<p>" +
          escapeHtml(result) +
          "</p>" +
          "<a href='" +
          "${CONFIG.WEBAPP_URL}?uebersicht=1" +
          "' target='_blank'>Zur Belegübersicht</a>" +
          "</div>";

      }
    )

    .withFailureHandler(
      function(error) {

        button.disabled =
          false;

        button.innerText =
          "🗑️ Beleg löschen";

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .deleteBelegByRow(
      rowNumber
    );

}


function chairmanApprovedManually() {

  if (!confirm("Eintragen, dass der 1. Vorsitzende diesen Beleg genehmigt hat?")) {
    return;
  }

  const button =
    document.getElementById(
      "chairmanApprovedButton"
    );

  button.disabled = true;

  google.script.run

    .withSuccessHandler(
      function(result) {

        button.style.display = "none";

        showMessage(
          result
        );

      }
    )

    .withFailureHandler(
      function(error) {

        button.disabled = false;

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .chairmanApprovedManuallyByRow(
      rowNumber
    );

}


function chooseAction(
  action
) {

  if (action === "Zahlung an 2. Kassierer") {
    sendToSecondCashier();
    return;
  }

  const buttons =
    document.querySelectorAll(
      ".decision-button"
    );

  buttons.forEach(
    function(button) {
      button.disabled = true;
    }
  );


  google.script.run

    .withSuccessHandler(
      function(result) {

        showMessage(
          result
        );

        buttons.forEach(
          function(button) {
            button.disabled = false;
          }
        );

      }
    )

    .withFailureHandler(
      function(error) {

        buttons.forEach(
          function(button) {
            button.disabled = false;
          }
        );

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .entscheidungTreffenByRow(
      rowNumber,
      action
    );

}




${formDataEditorJs()}
</script>

</head>

<body>

<div class="container">

<h1>
Belegprüfung KGV Neuenhof
</h1>

<div class="beleg-id">
Beleg ${htmlEscape(result.values[COL.BELEG_ID - 1] || belegId)}
</div>


<div
  class="status-box ${getStatusClass(status)}"
>
Status:
${htmlEscape(status)}
</div>


<table style="margin-top:20px">

<tr>
<td class="label">Zeitstempel</td>
<td>${htmlEscape(timestamp)}</td>
</tr>

${vorgangRowsHtml(row)}

<tr>
<td class="label">Einreicher</td>
<td id="cell_einreicher">${htmlEscape(einreicher)}</td>
</tr>

<tr>
<td class="label">Was wurde gekauft?</td>
<td id="cell_gekauft">${htmlEscape(gekauft)}</td>
</tr>

<tr>
<td class="label">Verwendungszweck</td>
<td id="cell_zweck">${htmlEscape(zweck)}</td>
</tr>

<tr>
<td class="label">Lieferant</td>
<td id="cell_lieferant">${htmlEscape(lieferant)}</td>
</tr>

<tr>
<td class="label">Betrag</td>
<td id="cell_betrag">${htmlEscape(formatAmountDisplay(betrag))} €</td>
</tr>

<tr>
<td class="label">Auszahlung an</td>
<td id="cell_auszahlungAn">${htmlEscape(auszahlungAn)}</td>
</tr>

<tr>
<td class="label">${htmlEscape(getColumnHeader(COL.RECHNUNGSNUMMER, "Rechnungsnummer"))}</td>
<td>${htmlEscape(rechnungsnummer)}</td>
</tr>

<tr>
<td class="label">E-Mail</td>
<td>${htmlEscape(email)}</td>
</tr>

<tr>
<td class="label">Dateiname</td>
<td id="cell_dateiname">${htmlEscape(dateiname)}</td>
</tr>

</table>

${formDataEditorHtml(row)}



${
  upload
    ? `
<a
  class="button document-button"
  href="${htmlEscape(upload)}"
  target="_blank"
>
📄 Rechnung / Originalbeleg öffnen
</a>
`
    : ""
}


<div class="section">

<div class="section-title">
Originalbeleg
</div>

<div id="originalStatus">

${originalAnzeige}

${
  originalDatum
    ? "<br>Datum: " +
      htmlEscape(originalDatum)
    : ""
}

</div>


${
  originalbeleg !== "Ja"
    ? `
<button
  id="originalButton"
  class="button original-button"
  onclick="originalbelegVorhanden()"
>
✓ Originalbeleg vorhanden
</button>
`
    : ""
}

</div>


<div class="section ${einnahme ? "hidden" : ""}">

<div class="section-title">
Zahlungsdaten
</div>

<div class="form-row">

<div class="form-label">
Zahlungsempfänger
</div>

<div class="payee-wrapper">

<input
  id="zahlungsempfaenger"
  value="${htmlEscape(zahlungsempfaenger)}"
  autocomplete="off"
  oninput="updatePayeeSuggestions()"
  placeholder="Name eingeben"
/>

<div
  id="payeeSuggestions"
  class="payee-suggestions"
></div>

</div>

</div>


<div class="form-row">

<div class="form-label">
IBAN
</div>

<input
  id="iban"
  value="${htmlEscape(iban)}"
/>

</div>


<div class="form-row">

<div class="form-label">
BIC
</div>

<input
  id="bic"
  value="${htmlEscape(bic)}"
/>

</div>


<div class="form-row">

<div class="form-label">
Zahlungsverwendungszweck
</div>

<input
  id="zahlungszweck"
  value="${htmlEscape(zahlungszweck)}"
/>

</div>


<div class="form-row">

<div class="form-label">
Information zur Überweisung
</div>

<textarea
  id="info"
>${htmlEscape(infoUeberweisung)}</textarea>

</div>

<div class="section">
<div class="section-title">SEPA-QR-Code</div>
<div id="sepaQrCode" style="text-align:center;padding:10px">
  <div class="small">SEPA-QR-Code wird hier angezeigt.</div>
</div>
</div>


<div class="checkbox-row">

<label>

<input
  type="checkbox"
  id="masterUpdate"
>

Zahlungsempfänger in den Stammdaten
speichern/aktualisieren

</label>

</div>


<button
  id="savePaymentButton"
  class="button save-button"
  onclick="savePaymentData()"
>
💾 Zahlungsdaten speichern
</button>


<div
  style="margin-top:15px"
>

<label>

<input
  type="checkbox"
  id="originalConfirmed"
  ${
    originalbeleg === "Ja"
      ? "checked"
      : ""
  }
/>

Originalbeleg ist vorhanden

</label>

</div>


<button
  id="sendPaymentButton"
  class="button payment-button"
  onclick="sendToSecondCashier()"
>
→ An 2. Kassierer zur Überweisung
</button>

</div>


<div class="section">

<div class="section-title">
Genehmigung / Bearbeitung
</div>

<div id="actionMessage"
class="message"
style="display:none">
</div>


<div>
${entscheidungsAnzeige}
</div>


${
  genehmigung
    ? `
<div style="margin-top:10px">
<strong>Vorsitzender:</strong>
${htmlEscape(genehmigung)}
<br>
Datum:
${htmlEscape(genehmigungDatum)}

${
  kommentarVorsitzender
    ? "<br><br><strong>Kommentar:</strong><br>" +
      htmlEscape(kommentarVorsitzender)
    : ""
}

</div>
`
    : ""
}


<div class="warning">

Die Bearbeitungsbuttons bleiben für den 1. Kassierer
auch bei laufender Genehmigung oder Überweisung sichtbar.

</div>


${
  einnahme
    ? `
<div class="message">
Einnahme: keine Genehmigung und keine Überweisung nötig.
Bitte unten den Zahlungseingang und die Buchung eintragen.
</div>
`
    : `
<button
  class="button decision-button"
  onclick="chooseAction('Genehmigung Vorsitzender')"
>
→ An 1. Vorsitzenden zur Genehmigung
</button>

${
  genehmigung
    ? ""
    : `
<button
  id="chairmanApprovedButton"
  class="button approve-button"
  onclick="chairmanApprovedManually()"
>
✓ 1. Vorsitzender hat genehmigt
</button>
`
}

<button
  class="button decision-button"
  onclick="chooseAction('Zahlung an 2. Kassierer')"
>
→ Direkt an 2. Kassierer zur Überweisung
</button>
`
}


<button
  class="button decision-button"
  onclick="chooseAction('Ablage')"
>
→ Nur zur Ablage
</button>


<button
  id="deleteButton"
  class="button delete-button"
  onclick="deleteBeleg()"
>
🗑️ Beleg löschen
</button>

</div>


<div class="section">

<div class="section-title">
Überweisung / Nachbearbeitung
</div>


<table>

<tr>
<td class="label">
An 2. Kassierer gesendet
</td>
<td>
${htmlEscape(an2Kassierer)}
</td>
</tr>

<tr>
<td class="label">
Überweisung erfolgt am
</td>
<td>
${htmlEscape(ueberwiesenAm)}
</td>
</tr>

<tr>
<td class="label">
Kommentar 2. Kassierer
</td>
<td>
${htmlEscape(kommentar2Kassierer)}
</td>
</tr>

<tr>
<td class="label">
${bezahltLabel}
</td>
<td>
${htmlEscape(bezahltAm)}
</td>
</tr>

<tr>
<td class="label">
Gebucht am
</td>
<td>
${htmlEscape(gebuchtAm)}
</td>
</tr>

</table>


<div class="form-row">

<div class="form-label">
${bezahltLabel}
</div>

<input
  type="date"
  id="bezahltDatum"
  value="${htmlDateValue(row[COL.BEZAHLT_AM - 1])}"
>

</div>


<div class="form-row">

<div class="form-label">
Gebucht am
</div>

<input
  type="date"
  id="gebuchtDatum"
  value="${htmlDateValue(row[COL.GEBUCHT_AM - 1])}"
>

</div>


<button
  class="button save-button"
  onclick="editAccountingDates()"
>
💾 Datumsänderungen speichern
</button>


<button
  class="button payment-button"
  onclick="markAsPaid()"
>
${einnahme ? "✓ Zahlung eingegangen" : "✓ Als bezahlt markieren"}
</button>


<button
  class="button payment-button"
  onclick="markAsBooked()"
>
✓ Als gebucht markieren
</button>

</div>


<a
  class="button back-button"
  href="${CONFIG.WEBAPP_URL}?uebersicht=1"
  target="_blank"
>
← Zur Belegübersicht
</a>


</div>

</body>

</html>

`;

}


/************************************************************
 * TABELLENZEILEN VORGANGSART / ZUSATZANGABEN
 ************************************************************/


function vorgangRowsHtml(values) {

  const zusatz =
    String(values[COL.ZUSATZANGABEN - 1] || "").trim();

  return (
    "<tr><td class=\"label\">Vorgangsart</td><td>" +
    htmlEscape(getVorgangsart(values)) +
    (isEinnahme(values) ? " <strong>(Einnahme)</strong>" : "") +
    "</td></tr>" +
    (zusatz
      ? "<tr><td class=\"label\">Zusatzangaben</td><td>" +
        htmlEscape(zusatz).split("\n").join("<br>") +
        "</td></tr>"
      : "")
  );

}


/************************************************************
 * DATUM FÜR INPUT TYPE=DATE
 ************************************************************/


function htmlDateValue(
  value
) {

  if (!value) {
    return "";
  }

  try {

    return Utilities.formatDate(
      new Date(value),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

  } catch (e) {

    return "";

  }

}


/************************************************************
 * ORIGINALBELEG
 ************************************************************/



/************************************************************
 * DIREKTE BUTTON-AKTIONEN ÜBER TABELLENZEILE
 *
 * Die Detailseite übergibt die echte Tabellenzeile.
 * Dadurch entfällt die komplette ID/UUID-Auflösung im Browser.
 ************************************************************/

function getBelegIdFromRow(rowNumber) {
  const sheet = getSheet();
  const n = Number(rowNumber);

  if (!Number.isInteger(n) || n < 2 || n > sheet.getLastRow()) {
    throw new Error("Ungültige Belegzeile: " + rowNumber);
  }

  const id = String(
    sheet.getRange(n, COL.BELEG_ID).getValue() || ""
  ).trim();

  if (!id) {
    throw new Error("In dieser Zeile ist keine Beleg-ID vorhanden.");
  }

  return id;
}

function originalbelegVorhandenByRow(rowNumber) {
  return originalbelegVorhanden(getBelegIdFromRow(rowNumber));
}

function savePaymentDataByRow(rowNumber, data) {
  return savePaymentData(getBelegIdFromRow(rowNumber), data || {});
}

function sendToSecondCashierByRow(rowNumber, data) {
  const belegId = getBelegIdFromRow(rowNumber);

  // Die im Formular eingetragenen Zahlungsdaten zuerst speichern.
  // Sonst prüft sendToSecondCashier() die alten Werte aus der Tabelle
  // und meldet "kein Zahlungsempfänger", obwohl er im Formular steht.
  if (data) {
    savePaymentData(belegId, data);

    // Häkchen "Originalbeleg ist vorhanden" ebenfalls übernehmen.
    if (data.originalConfirmed) {
      const sheet = getSheet();
      const n = Number(rowNumber);
      if (sheet.getRange(n, COL.ORIGINALBELEG).getValue() !== "Ja") {
        originalbelegVorhanden(belegId);
      }
    }
  }

  return sendToSecondCashier(belegId);
}

function entscheidungTreffenByRow(rowNumber, action) {
  return entscheidungTreffen(getBelegIdFromRow(rowNumber), String(action || ""));
}

function chairmanDecisionByRow(rowNumber, decision, comment) {
  return chairmanDecision(
    getBelegIdFromRow(rowNumber),
    decision,
    comment || ""
  );
}

// Genehmigung, die außerhalb des Systems erfolgt ist (z. B. mündlich
// oder in einer Sitzung). Keine Mail an den Vorsitzenden.
function chairmanApprovedManuallyByRow(rowNumber) {

  const sheet = getSheet();
  const n = Number(rowNumber);

  getBelegIdFromRow(n);

  sheet.getRange(n, COL.GENEHMIGUNG_VORSITZENDER).setValue("Genehmigt");
  sheet.getRange(n, COL.DATUM_GENEHMIGUNG).setValue(new Date());
  sheet.getRange(n, COL.KOMMENTAR_VORSITZENDER).setValue("Genehmigung durch 1. Kassierer eingetragen");
  sheet.getRange(n, COL.STATUS).setValue("Genehmigt");
  sheet.getRange(n, COL.LETZTE_BEARBEITUNG).setValue("Genehmigung Vorsitzender eingetragen");

  return "Genehmigung des 1. Vorsitzenden wurde eingetragen.";

}

function transferDoneByRow(rowNumber, comment) {
  return transferDone(
    getBelegIdFromRow(rowNumber),
    comment || ""
  );
}

function markAsPaidByRow(rowNumber, dateString) {
  return markAsPaid(
    getBelegIdFromRow(rowNumber),
    String(dateString || "")
  );
}

function markAsBookedByRow(rowNumber, dateString) {
  return markAsBooked(
    getBelegIdFromRow(rowNumber),
    String(dateString || "")
  );
}

function setAccountingDatesByRow(rowNumber, data) {
  const payload = data || {};
  return setAccountingDates(
    getBelegIdFromRow(rowNumber),
    String(payload.bezahlt || ""),
    String(payload.gebucht || "")
  );
}

function deleteBelegByRow(rowNumber) {
  return deleteBeleg(getBelegIdFromRow(rowNumber));
}

function performButtonAction(rowNumber, action, payload) {
  const sheet = getSheet();
  const n = Number(rowNumber);
  if (!Number.isInteger(n) || n < 2 || n > sheet.getLastRow()) {
    throw new Error("Ungültige Belegzeile: " + rowNumber);
  }

  const actualBelegId = String(sheet.getRange(n, COL.BELEG_ID).getValue() || "").trim();
  if (!actualBelegId) throw new Error("In dieser Zeile ist keine Beleg-ID vorhanden.");

  switch (String(action || "")) {
    case "savePaymentData":
      return savePaymentData(actualBelegId, payload || {});
    case "sendToSecondCashier":
      return sendToSecondCashier(actualBelegId);
    case "entscheidungTreffen":
      return entscheidungTreffen(actualBelegId, String(payload || ""));
    case "chairmanDecision":
      return chairmanDecision(actualBelegId, payload && payload.decision, payload && payload.comment);
    case "transferDone":
      return transferDone(actualBelegId, String(payload || ""));
    case "markAsPaid":
      return markAsPaid(actualBelegId, String(payload || ""));
    case "markAsBooked":
      return markAsBooked(actualBelegId, String(payload || ""));
    case "setAccountingDates":
      return setAccountingDates(actualBelegId, payload && payload.bezahlt, payload && payload.gebucht);
    case "deleteBeleg":
      return deleteBeleg(actualBelegId);
    default:
      throw new Error("Unbekannte Aktion: " + action);
  }
}


function originalbelegVorhanden(
  belegId
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  sheet
    .getRange(
      row,
      COL.ORIGINALBELEG
    )
    .setValue(
      "Ja"
    );

  sheet
    .getRange(
      row,
      COL.DATUM_ORIGINALBELEG
    )
    .setValue(
      new Date()
    );


  return "Originalbeleg wurde bestätigt.";

}


/************************************************************
 * ENTSCHEIDUNG 1. KASSIERER
 ************************************************************/


function entscheidungTreffen(
  belegId,
  entscheidung
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  if (
    entscheidung ===
    "Genehmigung Vorsitzender"
  ) {

    sheet
      .getRange(
        row,
        COL.ENTSCHEIDUNG_KASSIERER
      )
      .setValue(
        "Genehmigung Vorsitzender"
      );

    sheet
      .getRange(
        row,
        COL.DATUM_ENTSCHEIDUNG
      )
      .setValue(
        new Date()
      );

    sheet
      .getRange(
        row,
        COL.STATUS
      )
      .setValue(
        "Genehmigung Vorsitzender"
      );

    sheet
      .getRange(
        row,
        COL.LETZTE_BEARBEITUNG
      )
      .setValue(
        "An 1. Vorsitzenden gesendet"
      );


    sendChairmanApprovalMail(
      belegId
    );


    return (
      "Beleg wurde an den 1. Vorsitzenden " +
      "zur Genehmigung gesendet."
    );

  }


  if (
    entscheidung ===
    "Zahlung an 2. Kassierer"
  ) {

    return sendToSecondCashier(
      belegId
    );

  }


  if (
    entscheidung ===
    "Ablage"
  ) {

    sheet
      .getRange(
        row,
        COL.ENTSCHEIDUNG_KASSIERER
      )
      .setValue(
        "Ablage"
      );

    sheet
      .getRange(
        row,
        COL.DATUM_ENTSCHEIDUNG
      )
      .setValue(
        new Date()
      );

    sheet
      .getRange(
        row,
        COL.STATUS
      )
      .setValue(
        "Ablage"
      );

    sheet
      .getRange(
        row,
        COL.LETZTE_BEARBEITUNG
      )
      .setValue(
        "Beleg zur Ablage"
      );


    return (
      "Beleg wurde zur Ablage markiert."
    );

  }


  throw new Error(
    "Unbekannte Entscheidung."
  );

}


/************************************************************
 * ZAHLUNGSDATEN SPEICHERN
 ************************************************************/


function savePaymentData(
  belegId,
  data
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  const recipient =
    String(
      data.zahlungsempfaenger || ""
    ).trim();

  const iban =
    String(
      data.iban || ""
    ).trim();

  const bic =
    String(
      data.bic || ""
    ).trim();

  const purpose =
    String(
      data.zahlungszweck || ""
    ).trim();

  const info =
    String(
      data.info || ""
    ).trim();


  sheet
    .getRange(
      row,
      COL.ZAHLUNGSEMPFAENGER
    )
    .setValue(
      recipient
    );

  sheet
    .getRange(
      row,
      COL.IBAN
    )
    .setValue(
      iban
    );

  sheet
    .getRange(
      row,
      COL.BIC
    )
    .setValue(
      bic
    );

  sheet
    .getRange(
      row,
      COL.ZAHLUNGSZWECK
    )
    .setValue(
      purpose
    );

  sheet
    .getRange(
      row,
      COL.INFO_UEBERWEISUNG
    )
    .setValue(
      info
    );

  sheet
    .getRange(
      row,
      COL.ZAHLUNGSDATEN_GESPEICHERT
    )
    .setValue(
      new Date()
    );


  sheet
    .getRange(
      row,
      COL.LETZTE_BEARBEITUNG
    )
    .setValue(
      "Zahlungsdaten bearbeitet"
    );


  if (
    data.masterUpdate &&
    recipient
  ) {

    savePayeeToMasterData(
      recipient,
      iban,
      bic
    );

  }


  return (
    "Zahlungsdaten wurden gespeichert."
  );

}


/************************************************************
 * ZAHLUNGSEMPFÄNGER – STAMMDATEN
 ************************************************************/


function getPaymentRecipients() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      CONFIG.PAYEE_SHEET_NAME
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        CONFIG.PAYEE_SHEET_NAME
      );

    sheet
      .getRange(1,1,1,4)
      .setValues([
        [
          "Name",
          "IBAN",
          "BIC",
          "aktiv/inaktiv"
        ]
      ]);

    return [];

  }


  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        4
      )
      .getValues();


  return values
    .filter(
      function(row) {

        const name =
          String(
            row[0] || ""
          ).trim();

        const active =
          String(
            row[3] || "aktiv"
          )
          .toLowerCase();

        return (
          name &&
          active !== "inaktiv"
        );

      }
    )
    .map(
      function(row) {

        return {

          name:
            String(row[0] || ""),

          iban:
            String(row[1] || ""),

          bic:
            String(row[2] || "")

        };

      }
    );

}


/************************************************************
 * STAMMDATEN SPEICHERN / AKTUALISIEREN
 ************************************************************/


function savePayeeToMasterData(
  name,
  iban,
  bic
) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      CONFIG.PAYEE_SHEET_NAME
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        CONFIG.PAYEE_SHEET_NAME
      );

    sheet
      .getRange(1,1,1,4)
      .setValues([
        [
          "Name",
          "IBAN",
          "BIC",
          "aktiv/inaktiv"
        ]
      ]);

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow >= 2) {

    const values =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          4
        )
        .getValues();


    for (
      let i = 0;
      i < values.length;
      i++
    ) {

      if (
        String(values[i][0])
          .trim()
          .toLowerCase() ===
        String(name)
          .trim()
          .toLowerCase()
      ) {

        sheet
          .getRange(
            i + 2,
            1,
            1,
            3
          )
          .setValues([
            [
              name,
              iban,
              bic
            ]
          ]);

        return;

      }

    }

  }


  sheet
    .appendRow([
      name,
      iban,
      bic,
      "aktiv"
    ]);

}


/************************************************************
 * 2. KASSIERER ÜBERGABE
 ************************************************************/


function sendToSecondCashier(
  belegId
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  const data =
    sheet
      .getRange(
        row,
        1,
        1,
        COL.INTERNE_ID
      )
      .getValues()[0];


  const originalbeleg =
    data[
      COL.ORIGINALBELEG - 1
    ];

  const recipient =
    data[
      COL.ZAHLUNGSEMPFAENGER - 1
    ];

  const iban =
    data[
      COL.IBAN - 1
    ];

  const purpose =
    data[
      COL.ZAHLUNGSZWECK - 1
    ];


  if (
    originalbeleg !== "Ja"
  ) {

    throw new Error(
      "Der Originalbeleg wurde noch nicht bestätigt."
    );

  }


  if (!recipient) {

    throw new Error(
      "Es wurde kein Zahlungsempfänger angegeben."
    );

  }


  if (!iban) {

    throw new Error(
      "Es wurde keine IBAN angegeben."
    );

  }


  if (!purpose) {

    throw new Error(
      "Es wurde kein Zahlungsverwendungszweck angegeben."
    );

  }


  sheet
    .getRange(
      row,
      COL.STATUS
    )
    .setValue(
      "Zahlung vorbereitet"
    );


  sheet
    .getRange(
      row,
      COL.AN_2_KASSIERER_GESENDET
    )
    .setValue(
      new Date()
    );


  sheet
    .getRange(
      row,
      COL.LETZTE_BEARBEITUNG
    )
    .setValue(
      "An 2. Kassierer gesendet"
    );


  const rechnungsnummer =
    data[
      COL.RECHNUNGSNUMMER - 1
    ];


  const subject =
    "Bitte neue Überweisung anlegen – " +
    belegId;


  const url =
    makeWebUrl(
      belegId,
      {
        rolle:
          "zweiter_kassierer"
      }
    );


  const body =
    "Eine neue Überweisung wurde vorbereitet.\n\n" +

    "Beleg-ID: " +
    belegId +
    "\n" +

    getColumnHeader(COL.RECHNUNGSNUMMER, "Rechnungsnummer") + ": " +
    (
      rechnungsnummer || "-"
    ) +
    "\n\n" +

    "Bitte die Überweisung bearbeiten:\n" +
    url;


  sendMail(
    CONFIG.ZWEITER_KASSIERER_EMAIL,
    subject,
    body
  );


  return (
    "Beleg wurde an den 2. Kassierer " +
    "zur Überweisung übergeben."
  );

}


/************************************************************
 * 1. VORSITZENDER – E-MAIL
 ************************************************************/


function sendChairmanApprovalMail(
  belegId
) {

  const result =
    getRowData(
      belegId
    );

  if (!result) {
    return;
  }


  const row =
    result.values;


  const supplier =
    row[
      COL.LIEFERANT - 1
    ];

  const amount =
    row[
      COL.BETRAG - 1
    ];

  const rechnungsnummer =
    row[
      COL.RECHNUNGSNUMMER - 1
    ];


  const url =
    makeWebUrl(
      belegId,
      {
        rolle:
          "vorsitzender"
      }
    );


  const subject =
    "Beleg zur Genehmigung – " +
    belegId;


  const body =
    "Ein Beleg wartet auf Ihre Genehmigung.\n\n" +

    "Beleg-ID: " +
    belegId +
    "\n" +

    getColumnHeader(COL.RECHNUNGSNUMMER, "Rechnungsnummer") + ": " +
    (
      rechnungsnummer || "-"
    ) +
    "\n" +

    "Lieferant: " +
    supplier +
    "\n" +

    "Betrag: " +
    amount +
    " €\n\n" +

    "Beleg prüfen und entscheiden:\n" +
    url;


  sendMail(
    CONFIG.VORSITZENDER_EMAIL,
    subject,
    body
  );

}


/************************************************************
 * VORSITZENDER – SEITE
 ************************************************************/


function createChairmanPage(
  belegId, useInternalId
) {

  const result =
    getRowData(belegId, useInternalId);

  if (!result) {

    return errorPage(
      "Der Beleg wurde nicht gefunden."
    );

  }


  const row =
    result.values;


  const status =
    row[
      COL.STATUS - 1
    ];

  const einreicher =
    row[
      COL.EINREICHER - 1
    ];

  const gekauft =
    row[
      COL.GEKAUFT - 1
    ];

  const zweck =
    row[
      COL.ZWECK - 1
    ];

  const lieferant =
    row[
      COL.LIEFERANT - 1
    ];

  const betrag =
    row[
      COL.BETRAG - 1
    ];

  const auszahlungAn =
    row[
      COL.AUSZAHLUNG_AN - 1
    ];

  const rechnungsnummer =
    row[
      COL.RECHNUNGSNUMMER - 1
    ];

  const email =
    row[
      COL.EMAIL - 1
    ];

  const upload =
    row[
      COL.UPLOAD - 1
    ];

  const filename =
    row[
      COL.DATEINAME - 1
    ];

  const genehmigung =
    row[
      COL.GENEHMIGUNG_VORSITZENDER - 1
    ];

  const genehmigungDatum =
    formatDate(
      row[
        COL.DATUM_GENEHMIGUNG - 1
      ]
    );

  const kommentar =
    row[
      COL.KOMMENTAR_VORSITZENDER - 1
    ];


  return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Beleggenehmigung KGV Neuenhof</title>

<style>

${commonCss()}

</style>

<script>

const belegId =
  ${jsonForHtml(String(result.values[COL.BELEG_ID - 1] || belegId))};

const rowNumber =
  ${jsonForHtml(result.rowNumber)};


function chairmanDecision(
  decision
) {

  const comment =
    document.getElementById(
      "chairmanComment"
    ).value;


  const buttons =
    document.querySelectorAll(
      ".chairman-button"
    );

  buttons.forEach(
    function(button) {
      button.disabled = true;
    }
  );


  google.script.run

    .withSuccessHandler(
      function(result) {

        const box =
          document.getElementById(
            "result"
          );

        box.innerHTML =
          "✅ " +
          result;

        box.style.display =
          "block";

        buttons.forEach(
          function(button) {
            button.style.display = "none";
          }
        );

      }
    )

    .withFailureHandler(
      function(error) {

        buttons.forEach(
          function(button) {
            button.disabled = false;
          }
        );

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .chairmanDecisionByRow(
      rowNumber,
      decision,
      comment
    );

}


${formDataEditorJs()}
</script>

</head>

<body>

<div class="container">

<h1>
Beleggenehmigung KGV Neuenhof
</h1>

<div class="beleg-id">
Beleg ${htmlEscape(result.values[COL.BELEG_ID - 1] || belegId)}
</div>


<div
  class="status-box ${getStatusClass(status)}"
>
Status:
${htmlEscape(status)}
</div>


<table style="margin-top:20px">

${vorgangRowsHtml(row)}

<tr>
<td class="label">Einreicher</td>
<td id="cell_einreicher">${htmlEscape(einreicher)}</td>
</tr>

<tr>
<td class="label">Was wurde gekauft?</td>
<td id="cell_gekauft">${htmlEscape(gekauft)}</td>
</tr>

<tr>
<td class="label">Verwendungszweck</td>
<td id="cell_zweck">${htmlEscape(zweck)}</td>
</tr>

<tr>
<td class="label">Lieferant</td>
<td id="cell_lieferant">${htmlEscape(lieferant)}</td>
</tr>

<tr>
<td class="label">Betrag</td>
<td id="cell_betrag">${htmlEscape(formatAmountDisplay(betrag))} €</td>
</tr>

<tr>
<td class="label">Auszahlung an</td>
<td id="cell_auszahlungAn">${htmlEscape(auszahlungAn)}</td>
</tr>

<tr>
<td class="label">${htmlEscape(getColumnHeader(COL.RECHNUNGSNUMMER, "Rechnungsnummer"))}</td>
<td>${htmlEscape(rechnungsnummer)}</td>
</tr>

<tr>
<td class="label">E-Mail</td>
<td>${htmlEscape(email)}</td>
</tr>

<tr>
<td class="label">Dateiname</td>
<td id="cell_dateiname">${htmlEscape(filename)}</td>
</tr>

</table>

${formDataEditorHtml(row)}



${
  upload
    ? `
<a
  class="button document-button"
  href="${htmlEscape(upload)}"
  target="_blank"
>
📄 Rechnung / Originalbeleg öffnen
</a>
`
    : ""
}


<div class="section">

<div class="section-title">
Entscheidung des 1. Vorsitzenden
</div>


<div class="form-row">

<div class="form-label">
Kommentar
</div>

<textarea
  id="chairmanComment"
  placeholder="Optionaler Kommentar"
>${htmlEscape(kommentar)}</textarea>

</div>


${
  genehmigung
    ? `
<div class="message">

<strong>
Bereits entschieden:
</strong>

<br>

${htmlEscape(genehmigung)}

<br>

Datum:
${htmlEscape(genehmigungDatum)}

</div>
`
    : `
<button
  class="button approve-button chairman-button"
  onclick="chairmanDecision('Genehmigt')"
>
✅ Beleg genehmigen
</button>


<button
  class="button reject-button chairman-button"
  onclick="chairmanDecision('Abgelehnt')"
>
❌ Beleg ablehnen
</button>
`
}


<div id="result"
class="message"
style="display:none">
</div>

</div>


<a
  class="button back-button"
  href="${CONFIG.WEBAPP_URL}?uebersicht=1&rolle=vorsitzender"
  target="_blank"
>
← Meine Genehmigungen
</a>


</div>

</body>

</html>

`;

}


/************************************************************
 * VORSITZENDER – ENTSCHEIDUNG
 ************************************************************/


function chairmanDecision(
  belegId,
  decision,
  comment
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  if (
    decision !== "Genehmigt" &&
    decision !== "Abgelehnt"
  ) {

    throw new Error(
      "Ungültige Entscheidung."
    );

  }


  sheet
    .getRange(
      row,
      COL.GENEHMIGUNG_VORSITZENDER
    )
    .setValue(
      decision
    );


  sheet
    .getRange(
      row,
      COL.DATUM_GENEHMIGUNG
    )
    .setValue(
      new Date()
    );


  sheet
    .getRange(
      row,
      COL.KOMMENTAR_VORSITZENDER
    )
    .setValue(
      comment || ""
    );


  sheet
    .getRange(
      row,
      COL.STATUS
    )
    .setValue(
      decision
    );


  sheet
    .getRange(
      row,
      COL.LETZTE_BEARBEITUNG
    )
    .setValue(
      "Vorsitzender: " +
      decision
    );


  const url =
    makeWebUrl(
      belegId
    );


  let body =
    "Der 1. Vorsitzende hat den Beleg entschieden.\n\n" +

    "Beleg-ID: " +
    belegId +
    "\n\n" +

    "Entscheidung: " +
    decision +
    "\n";


  if (comment) {

    body +=
      "\nKommentar:\n" +
      comment +
      "\n";

  }


  body +=
    "\nAktueller Beleg:\n" +
    url;


  sendMail(
    CONFIG.KASSIERER_EMAIL,
    "Entscheidung zum Beleg – " +
      belegId,
    body
  );


  return (
    "Beleg wurde als '" +
    decision +
    "' gespeichert."
  );

}


/************************************************************
 * 2. KASSIERER – ÜBERWEISUNGSSEITE
 ************************************************************/


function createPaymentPage(
  belegId, useInternalId
) {

  const result =
    getRowData(belegId, useInternalId);

  if (!result) {

    return errorPage(
      "Der Beleg wurde nicht gefunden."
    );

  }


  const row =
    result.values;


  const status =
    row[
      COL.STATUS - 1
    ];

  const einreicher =
    row[
      COL.EINREICHER - 1
    ];

  const gekauft =
    row[
      COL.GEKAUFT - 1
    ];

  const zweck =
    row[
      COL.ZWECK - 1
    ];

  const lieferant =
    row[
      COL.LIEFERANT - 1
    ];

  const betrag =
    row[
      COL.BETRAG - 1
    ];

  const auszahlungAn =
    row[
      COL.AUSZAHLUNG_AN - 1
    ];

  const rechnungsnummer =
    row[
      COL.RECHNUNGSNUMMER - 1
    ];

  const email =
    row[
      COL.EMAIL - 1
    ];

  const upload =
    row[
      COL.UPLOAD - 1
    ];

  const recipient =
    row[
      COL.ZAHLUNGSEMPFAENGER - 1
    ];

  const iban =
    row[
      COL.IBAN - 1
    ];

  const bic =
    row[
      COL.BIC - 1
    ];

  const paymentPurpose =
    row[
      COL.ZAHLUNGSZWECK - 1
    ];

  const info =
    row[
      COL.INFO_UEBERWEISUNG - 1
    ];

  const comment =
    row[
      COL.KOMMENTAR_2_KASSIERER - 1
    ];

  const ueberwiesenAm =
    formatDate(
      row[
        COL.UEBERWIESEN_AM - 1
      ]
    );


  return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Überweisung – KGV Neuenhof</title>

<style>

${commonCss()}

</style>

<script>
const sepaPaymentPayload =
  ${jsonForHtml(
    buildSepaQrPayload(recipient, iban, bic, betrag, paymentPurpose)
  )};

function renderPaymentQr() {
  const box = document.getElementById("sepaQrCodePayment");
  if (!box) return;

  const payload = String(sepaPaymentPayload || "");

  if (!payload) {
    box.innerHTML =
      "<div class='warning'>Kein QR-Code möglich: Zahlungsempfänger, " +
      "IBAN oder Verwendungszweck fehlen. Bitte beim 1. Kassierer nachfragen.</div>";
    return;
  }

  box.innerHTML =
    "<img alt='SEPA QR-Code' " +
    "style='max-width:320px;width:100%;height:auto' " +
    "src='https://quickchart.io/qr?text=" +
    encodeURIComponent(payload) +
    "&size=320'>" +
    "<div class='small' style='margin-top:8px'>SEPA-Überweisungsdaten</div>";
}

window.addEventListener("load", renderPaymentQr);



const belegId =
  ${jsonForHtml(String(result.values[COL.BELEG_ID - 1] || belegId))};

const rowNumber =
  ${jsonForHtml(result.rowNumber)};


function transferDone() {

  const comment =
    document.getElementById(
      "comment"
    ).value;


  const button =
    document.getElementById(
      "transferButton"
    );

  button.disabled =
    true;

  button.innerText =
    "Wird gespeichert...";


  google.script.run

    .withSuccessHandler(
      function(result) {

        document.getElementById(
          "result"
        ).style.display =
          "block";

        document.getElementById(
          "result"
        ).innerHTML =
          "✅ " +
          result;

        button.innerText =
          "✓ Überweisung erfolgt";

      }
    )

    .withFailureHandler(
      function(error) {

        button.disabled =
          false;

        button.innerText =
          "💶 Überweisung erfolgt";

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .transferDoneByRow(
      rowNumber,
      comment
    );

}

</script>

</head>

<body>

<div class="container">

<h1>
Überweisung
</h1>

<div class="beleg-id">
Beleg ${htmlEscape(result.values[COL.BELEG_ID - 1] || belegId)}
</div>


<div
  class="status-box ${getStatusClass(status)}"
>
Status:
${htmlEscape(status)}
</div>


<table style="margin-top:20px">

${vorgangRowsHtml(row)}

<tr>
<td class="label">Einreicher</td>
<td>${htmlEscape(einreicher)}</td>
</tr>

<tr>
<td class="label">Was wurde gekauft?</td>
<td>${htmlEscape(gekauft)}</td>
</tr>

<tr>
<td class="label">Verwendungszweck</td>
<td>${htmlEscape(zweck)}</td>
</tr>

<tr>
<td class="label">Lieferant</td>
<td>${htmlEscape(lieferant)}</td>
</tr>

<tr>
<td class="label">Betrag</td>
<td>${htmlEscape(betrag)} €</td>
</tr>

<tr>
<td class="label">Auszahlung an</td>
<td>${htmlEscape(auszahlungAn)}</td>
</tr>

<tr>
<td class="label">${htmlEscape(getColumnHeader(COL.RECHNUNGSNUMMER, "Rechnungsnummer"))}</td>
<td>${htmlEscape(rechnungsnummer)}</td>
</tr>

<tr>
<td class="label">E-Mail</td>
<td>${htmlEscape(email)}</td>
</tr>

</table>


${
  upload
    ? `
<a
  class="button document-button"
  href="${htmlEscape(upload)}"
  target="_blank"
>
📄 Rechnung / Originalbeleg öffnen
</a>
`
    : ""
}


<div class="section">

<div class="section-title">
Zahlungsdaten
</div>

<table>

<tr>
<td class="label">Zahlungsempfänger</td>
<td>${htmlEscape(recipient)}</td>
</tr>

<tr>
<td class="label">IBAN</td>
<td>${htmlEscape(iban)}</td>
</tr>

<tr>
<td class="label">BIC</td>
<td>${htmlEscape(bic)}</td>
</tr>

<tr>
<td class="label">Betrag</td>
<td>${htmlEscape(betrag)} €</td>
</tr>

<tr>
<td class="label">Zahlungsverwendungszweck</td>
<td>${htmlEscape(paymentPurpose)}</td>
</tr>

</table>


${
  info
    ? `
<div class="section">
<strong>
Information zur Überweisung
</strong>
<br><br>
${htmlEscape(info)}
</div>
`
    : ""
}


<div class="section">

<div class="section-title">
SEPA-Überweisung
</div>

<p>
Bitte die Überweisung anhand der oben angezeigten
Zahlungsdaten durchführen oder den QR-Code mit der
Banking-App scannen.
</p>

<div id="sepaQrCodePayment" style="text-align:center;padding:10px">
  <div class="small">SEPA-QR-Code wird geladen...</div>
</div>

</div>


<div class="form-row">

<div class="form-label">
Kommentar des 2. Kassierers
</div>

<textarea
  id="comment"
  placeholder="Optionaler Kommentar"
>${htmlEscape(comment)}</textarea>

</div>


<div id="result"
class="message"
style="display:none">
</div>


${
  status === "Zahlung vorbereitet"
    ? `
<button
  id="transferButton"
  class="button payment-button"
  onclick="transferDone()"
>
💶 Überweisung erfolgt
</button>
`
    : `
<div class="message">

Überweisung bereits verarbeitet.

<br>

Überwiesen am:
${htmlEscape(ueberwiesenAm)}

</div>
`
}


<a
  class="button back-button"
  href="${CONFIG.WEBAPP_URL}?uebersicht=1&rolle=zweiter_kassierer"
  target="_blank"
>
← Meine Überweisungen
</a>

</div>

</body>

</html>

`;

}


/************************************************************
 * SEPA-QR-CODE (EPC/GiroCode)
 *
 * Liefert "" wenn Pflichtangaben fehlen.
 ************************************************************/


function buildSepaQrPayload(recipient, iban, bic, amount, purpose) {

  const name = String(recipient || "").trim();
  const cleanIban = String(iban || "").replace(/\s+/g, "").toUpperCase();
  const cleanBic = String(bic || "").replace(/\s+/g, "").toUpperCase();
  const text = String(purpose || "").trim();

  if (!name || !cleanIban || !text) {
    return "";
  }

  // Betrag kann als Zahl (12.5) oder Text ("12,50 €") in der Tabelle stehen.
  let value =
    typeof amount === "number"
      ? amount
      : Number(String(amount || "").replace(/[^\d,.-]/g, "").replace(",", "."));

  const amountPart =
    value > 0 ? "EUR" + value.toFixed(2) : "";

  return [
    "BCD",
    "002",
    "1",
    "SCT",
    cleanBic,
    name.substring(0, 70),
    cleanIban,
    amountPart,
    "",
    "",
    text.substring(0, 140)
  ].join("\n");

}


/************************************************************
 * 2. KASSIERER – ÜBERWEISUNG ERFOLGT
 ************************************************************/


function transferDone(
  belegId,
  comment
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  sheet
    .getRange(
      row,
      COL.STATUS
    )
    .setValue(
      "Überwiesen"
    );


  sheet
    .getRange(
      row,
      COL.UEBERWIESEN_AM
    )
    .setValue(
      new Date()
    );


  sheet
    .getRange(
      row,
      COL.KOMMENTAR_2_KASSIERER
    )
    .setValue(
      comment || ""
    );


  sheet
    .getRange(
      row,
      COL.LETZTE_BEARBEITUNG
    )
    .setValue(
      "Überweisung durch 2. Kassierer erfolgt"
    );


  const url =
    makeWebUrl(
      belegId
    );


  const body =
    "Überweisung angelegt, bitte bei der Sparkasse freigeben.\n\n" +

    "Beleg-ID: " +
    belegId +
    "\n\n" +

    (
      comment
        ? "Kommentar des 2. Kassierers:\n" +
          comment +
          "\n\n"
        : ""
    ) +

    "Beleg öffnen:\n" +
    url;


  sendMail(
    CONFIG.KASSIERER_EMAIL,
    "Überweisung erfolgt – " +
      belegId,
    body
  );


  sendMail(
    CONFIG.VORSITZENDER_EMAIL,
    "Überweisung erfolgt – " +
      belegId,
    body
  );


  return (
    "Die Überweisung wurde angelegt. " +
    "1. Kassierer und 1. Vorsitzender wurden informiert. " +
    "Bitte bei der Sparkasse freigeben."
  );

}


/************************************************************
 * BEZAHLT
 ************************************************************/


function markAsPaid(
  belegId,
  dateString
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  const date =
    parseDateInput(
      dateString
    );


  sheet
    .getRange(
      row,
      COL.BEZAHLT_AM
    )
    .setValue(
      date
    )
    .setNumberFormat(
      "dd.MM.yyyy"
    );


  const einnahme =
    isEinnahme(
      sheet.getRange(row, 1, 1, COL_MAX).getValues()[0]
    );

  sheet
    .getRange(
      row,
      COL.STATUS
    )
    .setValue(
      einnahme ? STATUS_EINNAHME_EINGEGANGEN : "Bezahlt"
    );


  sheet
    .getRange(
      row,
      COL.LETZTE_BEARBEITUNG
    )
    .setValue(
      einnahme ? "Zahlungseingang eingetragen" : "Als bezahlt markiert"
    );


  return (
    (einnahme
      ? "Der Zahlungseingang wurde eingetragen. Eingangsdatum: "
      : "Der Beleg wurde als bezahlt markiert. Zahlungsdatum: ") +
    Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "dd.MM.yyyy"
    ) +
    syncBelegCopiesSafe(row)
  );

}


/************************************************************
 * GEBUCHT
 ************************************************************/


function markAsBooked(
  belegId,
  dateString
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  const date =
    parseDateInput(
      dateString
    );


  sheet
    .getRange(
      row,
      COL.GEBUCHT_AM
    )
    .setValue(
      date
    )
    .setNumberFormat(
      "dd.MM.yyyy"
    );


  sheet
    .getRange(
      row,
      COL.STATUS
    )
    .setValue(
      "Gebucht"
    );


  sheet
    .getRange(
      row,
      COL.LETZTE_BEARBEITUNG
    )
    .setValue(
      "Als gebucht markiert"
    );


  return (
    "Der Beleg wurde als gebucht markiert. " +
    "Buchungsdatum: " +
    Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "dd.MM.yyyy"
    ) +
    syncBelegCopiesSafe(row)
  );

}


/************************************************************
 * BEZAHLT / GEBUCHT MANUELL ÄNDERN
 ************************************************************/


function setAccountingDates(
  belegId,
  bezahltString,
  gebuchtString
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  if (bezahltString) {

    sheet
      .getRange(
        row,
        COL.BEZAHLT_AM
      )
      .setValue(
        parseDateInput(
          bezahltString
        )
      )
      .setNumberFormat(
        "dd.MM.yyyy"
      );

  } else {

    sheet
      .getRange(
        row,
        COL.BEZAHLT_AM
      )
      .clearContent();

  }


  if (gebuchtString) {

    sheet
      .getRange(
        row,
        COL.GEBUCHT_AM
      )
      .setValue(
        parseDateInput(
          gebuchtString
        )
      )
      .setNumberFormat(
        "dd.MM.yyyy"
      );

  } else {

    sheet
      .getRange(
        row,
        COL.GEBUCHT_AM
      )
      .clearContent();

  }


  sheet
    .getRange(
      row,
      COL.LETZTE_BEARBEITUNG
    )
    .setValue(
      "Bezahlt-/Gebucht-Datum manuell geändert"
    );


  return (
    "Die Datumsangaben wurden gespeichert." +
    syncBelegCopiesSafe(row)
  );

}


/************************************************************
 * DATUM PARSEN
 ************************************************************/


function parseDateInput(
  dateString
) {

  if (!dateString) {

    throw new Error(
      "Kein Datum angegeben."
    );

  }


  const parts =
    String(dateString)
      .split("-");


  if (
    parts.length !== 3
  ) {

    throw new Error(
      "Ungültiges Datum."
    );

  }


  const year =
    Number(parts[0]);

  const month =
    Number(parts[1]);

  const day =
    Number(parts[2]);


  const date =
    new Date(
      year,
      month - 1,
      day
    );


  if (
    isNaN(date.getTime())
  ) {

    throw new Error(
      "Ungültiges Datum."
    );

  }


  return date;

}


/************************************************************
 * BELEGKOPIEN IM ORDNER "BELEGEINGANG"
 *
 * Dateiname:
 *   <Eingangsdatum>_<Lieferant>_gezahlt am_<Datum>_gebucht am_<Datum>
 * Unbekannte Daten werden mit "xx" ausgefüllt.
 *
 * Ablage:
 *   noch nicht bezahlt  -> Belegeingang
 *   bezahlt             -> Belegeingang/Bezahlt
 *   gebucht             -> Belegeingang/Gebucht
 *   Beleg gelöscht      -> Belegeingang/Gelöscht
 *
 * Die Drive-IDs der Kopien stehen in Spalte BELEGKOPIE_ID.
 * Die Originaldatei aus dem Formular bleibt unverändert liegen.
 ************************************************************/


function getOrCreateFolder(parent, name) {

  const folders =
    parent.getFoldersByName(name);

  return folders.hasNext()
    ? folders.next()
    : parent.createFolder(name);

}


function getBelegFolder(subfolderName) {

  const spreadsheetFile =
    DriveApp.getFileById(
      SpreadsheetApp.getActiveSpreadsheet().getId()
    );

  const parents =
    spreadsheetFile.getParents();

  const base =
    parents.hasNext()
      ? parents.next()
      : DriveApp.getRootFolder();

  const eingang =
    getOrCreateFolder(base, CONFIG.BELEG_ORDNER);

  return subfolderName
    ? getOrCreateFolder(eingang, subfolderName)
    : eingang;

}


function extractDriveFileIds(uploadValue) {

  // Ein Upload-Feld kann mehrere Links enthalten (durch Komma getrennt).
  const ids = [];
  const pattern = /id=([\w-]+)/g;
  let match;

  while ((match = pattern.exec(String(uploadValue || ""))) !== null) {
    if (ids.indexOf(match[1]) < 0) {
      ids.push(match[1]);
    }
  }

  return ids;

}


function fileNameDate(value) {

  if (!value) {
    return "xx";
  }

  const date = new Date(value);

  if (isNaN(date.getTime())) {
    return "xx";
  }

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

}


function getFileExtension(name) {

  const match =
    String(name || "").match(/\.[A-Za-z0-9]{1,5}$/);

  return match ? match[0] : "";

}


function buildBelegCopyBaseName(values) {

  const supplier =
    String(values[COL.LIEFERANT - 1] || "Unbekannt")
      .trim()
      .replace(/[^\wäöüÄÖÜß-]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .substring(0, 60) || "Unbekannt";

  return (
    fileNameDate(values[COL.ZEITSTEMPEL - 1]) + "_" +
    supplier + "_" +
    (isEinnahme(values) ? "eingegangen am_" : "gezahlt am_") +
    fileNameDate(values[COL.BEZAHLT_AM - 1]) + "_" +
    "gebucht am_" + fileNameDate(values[COL.GEBUCHT_AM - 1])
  );

}


function syncBelegCopies(rowNumber, deleted) {

  const sheet = getSheet();

  const values =
    sheet
      .getRange(rowNumber, 1, 1, COL_MAX)
      .getValues()[0];

  let copyIds =
    String(values[COL.BELEGKOPIE_ID - 1] || "")
      .split(",")
      .map(function(id) { return id.trim(); })
      .filter(String);

  const baseName =
    buildBelegCopyBaseName(values);


  // Zielordner nach Bearbeitungsstand
  let folderLabel = CONFIG.BELEG_ORDNER;
  let target;

  if (deleted) {
    target = getBelegFolder(CONFIG.BELEG_ORDNER_GELOESCHT);
    folderLabel += "/" + CONFIG.BELEG_ORDNER_GELOESCHT;
  } else if (values[COL.GEBUCHT_AM - 1]) {
    target = getBelegFolder(CONFIG.BELEG_ORDNER_GEBUCHT);
    folderLabel += "/" + CONFIG.BELEG_ORDNER_GEBUCHT;
  } else if (values[COL.BEZAHLT_AM - 1]) {
    target = getBelegFolder(CONFIG.BELEG_ORDNER_BEZAHLT);
    folderLabel += "/" + CONFIG.BELEG_ORDNER_BEZAHLT;
  } else {
    target = getBelegFolder("");
  }


  // Noch keine Kopie vorhanden (neuer Beleg oder Beleg von vor
  // Einführung dieser Funktion): jetzt aus dem Upload anlegen.
  if (!copyIds.length) {

    const sourceIds =
      extractDriveFileIds(values[COL.UPLOAD - 1]);

    if (!sourceIds.length) {
      return "";
    }

    copyIds =
      sourceIds.map(function(sourceId, index) {

        const source = DriveApp.getFileById(sourceId);

        const name =
          baseName +
          (sourceIds.length > 1 ? "_" + (index + 1) : "") +
          getFileExtension(source.getName());

        return source.makeCopy(name, target).getId();

      });

    sheet
      .getRange(rowNumber, COL.BELEGKOPIE_ID)
      .setValue(copyIds.join(","));

    return folderLabel;

  }


  // Vorhandene Kopien umbenennen und verschieben
  copyIds.forEach(function(copyId, index) {

    const file = DriveApp.getFileById(copyId);

    file.setName(
      baseName +
      (copyIds.length > 1 ? "_" + (index + 1) : "") +
      getFileExtension(file.getName())
    );

    file.moveTo(target);

  });

  return folderLabel;

}


// Ein Fehler bei der Ablage darf die eigentliche Aktion
// (Formulareingang, bezahlt, gebucht) nicht abbrechen.
// Rückgabe: Zusatztext für die Erfolgsmeldung.
function syncBelegCopiesSafe(rowNumber, deleted) {

  try {

    const folder = syncBelegCopies(rowNumber, deleted);

    return folder
      ? " Belegkopie liegt in " + folder + "."
      : "";

  } catch (error) {

    Logger.log("Belegkopie: " + error);

    return (
      " Hinweis: Die Belegkopie konnte nicht abgelegt werden (" +
      (error && error.message ? error.message : error) +
      ")."
    );

  }

}


/************************************************************
 * BELEG LÖSCHEN
 ************************************************************/


function deleteBeleg(
  belegId
) {

  const sheet =
    getSheet();

  const row =
    findRowByIdentifier(
      belegId
    );

  if (!row) {
    throw new Error(
      "Beleg nicht gefunden."
    );
  }


  // Belegkopie vor dem Löschen der Zeile nach "Gelöscht" verschieben,
  // danach sind die Drive-IDs nicht mehr in der Tabelle.
  const ablage =
    syncBelegCopiesSafe(row, true);


  sheet.deleteRow(
    row
  );


  return (
    "Beleg " +
    belegId +
    " wurde aus der Belegliste gelöscht." +
    ablage
  );

}


/************************************************************
 * ZAHLUNGSEMPFÄNGER AUTOCOMPLETE
 ************************************************************/


function getPayeeSuggestions(
  searchText
) {

  const all =
    getPaymentRecipients();

  const query =
    String(
      searchText || ""
    )
    .trim()
    .toLowerCase();


  if (!query) {
    return all;
  }


  return all.filter(
    function(item) {

      return (
        item.name
          .toLowerCase()
          .indexOf(query) >= 0
      );

    }
  );

}


/************************************************************
 * ANGABEN KORRIGIEREN (1. Kassierer und 1. Vorsitzender)
 *
 * Bearbeitbar: Einreicher, Gekauft, Verwendungszweck,
 * Lieferant, Betrag, Auszahlung an.
 * Der Betrag ist gesperrt, sobald überwiesen wurde.
 * Bei geändertem Lieferanten werden Originaldatei(en) und
 * Belegkopie umbenannt.
 ************************************************************/


const EDIT_FIELDS = [
  { id: "einreicher", col: COL.EINREICHER, label: "Einreicher" },
  { id: "gekauft", col: COL.GEKAUFT, label: "Was wurde gekauft?" },
  { id: "zweck", col: COL.ZWECK, label: "Verwendungszweck" },
  { id: "lieferant", col: COL.LIEFERANT, label: "Lieferant" },
  { id: "betrag", col: COL.BETRAG, label: "Betrag (€)" },
  { id: "auszahlungAn", col: COL.AUSZAHLUNG_AN, label: "Auszahlung an" }
];


function formatAmountDisplay(value) {

  if (value === "" || value === null || value === undefined) {
    return "";
  }

  return typeof value === "number"
    ? value.toFixed(2).replace(".", ",")
    : String(value);

}


function formDataEditorHtml(values) {

  const locked = isAmountLocked(values);

  const inputs =
    EDIT_FIELDS.map(function(field) {

      const raw = values[field.col - 1];

      const value =
        field.id === "betrag"
          ? formatAmountDisplay(raw)
          : String(raw === null || raw === undefined ? "" : raw);

      const disabled =
        field.id === "betrag" && locked;

      return (
        "<div class=\"form-row\">" +
        "<div class=\"form-label\">" + htmlEscape(field.label) + "</div>" +
        "<input id=\"edit_" + field.id + "\" value=\"" + htmlEscape(value) + "\"" +
        (field.id === "betrag" ? " inputmode=\"decimal\"" : "") +
        (disabled ? " disabled" : "") + ">" +
        (disabled
          ? "<div class=\"small\" style=\"font-size:13px;color:#666\">" +
            "Der Betrag kann nach der Überweisung nicht mehr geändert werden.</div>"
          : "") +
        "</div>"
      );

    }).join("");

  return `
<button
  id="editButton"
  class="button secondary-button"
  onclick="toggleFormDataEdit(true)"
>
✏️ Angaben bearbeiten
</button>

<div id="editMessage" class="message" style="display:none"></div>

<div id="editSection" class="section hidden">

<div class="section-title">Angaben korrigieren</div>

${inputs}

<button
  id="saveFormDataButton"
  class="button save-button"
  onclick="saveFormData()"
>
💾 Korrektur speichern
</button>

<button
  class="button back-button"
  onclick="cancelFormDataEdit()"
>
Abbrechen
</button>

</div>
`;

}


// Browser-Skript für die Korrektur. Erwartet die Variable rowNumber.
// Keine Backslashes verwenden: der Text steht in einem Template-String.
function formDataEditorJs() {

  return `

const EDIT_IDS = ${jsonForHtml(EDIT_FIELDS.map(function(f) { return f.id; }))};


function toggleFormDataEdit(show) {

  document.getElementById("editSection").classList.toggle("hidden", !show);
  document.getElementById("editButton").classList.toggle("hidden", show);

  if (show) {
    document.getElementById("editMessage").style.display = "none";
  }

}


function cancelFormDataEdit() {

  EDIT_IDS.forEach(function(id) {
    const input = document.getElementById("edit_" + id);
    input.value = input.defaultValue;
  });

  toggleFormDataEdit(false);

}


function saveFormData() {

  const data = {};

  EDIT_IDS.forEach(function(id) {
    data[id] = document.getElementById("edit_" + id).value;
  });

  const button = document.getElementById("saveFormDataButton");
  button.disabled = true;
  button.innerText = "Wird gespeichert...";

  google.script.run

    .withSuccessHandler(function(result) {

      button.disabled = false;
      button.innerText = "💾 Korrektur speichern";

      Object.keys(result.values).forEach(function(id) {

        const cell = document.getElementById("cell_" + id);
        const input = document.getElementById("edit_" + id);
        const value = result.values[id];

        if (cell) {
          cell.innerText =
            id === "betrag" && value !== "" ? value + " €" : value;
        }

        if (input) {
          input.value = value;
          input.defaultValue = value;
        }

      });

      toggleFormDataEdit(false);

      const box = document.getElementById("editMessage");
      box.innerText = "✅ " + result.message;
      box.style.display = "block";

      if (typeof onFormDataSaved === "function") {
        onFormDataSaved(result);
      }

    })

    .withFailureHandler(function(error) {

      button.disabled = false;
      button.innerText = "💾 Korrektur speichern";
      alert("Fehler: " + (error && error.message ? error.message : error));

    })

    .saveFormDataByRow(rowNumber, data);

}

`;

}


function saveFormDataByRow(rowNumber, data) {

  data = data || {};

  const sheet = getSheet();
  const n = Number(rowNumber);
  const belegId = getBelegIdFromRow(n);
  const values = sheet.getRange(n, 1, 1, COL_MAX).getValues()[0];
  const text = function(value) { return String(value === null || value === undefined ? "" : value).trim(); };


  // Betrag prüfen
  const newAmount = parseAmount(data.betrag);
  const oldAmountRaw = values[COL.BETRAG - 1];
  let oldAmount;

  try {
    oldAmount = parseAmount(oldAmountRaw);
  } catch (e) {
    oldAmount = text(oldAmountRaw);
  }

  const amountChanged = newAmount !== oldAmount;

  if (amountChanged && isAmountLocked(values)) {
    throw new Error(
      "Der Betrag kann nach der Überweisung nicht mehr geändert werden."
    );
  }


  const oldSupplier = text(values[COL.LIEFERANT - 1]);
  const newSupplier = text(data.lieferant);

  EDIT_FIELDS.forEach(function(field) {

    if (field.id === "betrag") {
      if (amountChanged) {
        sheet.getRange(n, field.col).setValue(newAmount);
      }
      return;
    }

    sheet.getRange(n, field.col).setValue(text(data[field.id]));

  });

  sheet
    .getRange(n, COL.LETZTE_BEARBEITUNG)
    .setValue("Angaben korrigiert");


  let hint = "";
  let dateiname = text(values[COL.DATEINAME - 1]);

  if (newSupplier !== oldSupplier) {

    // Originaldatei(en) umbenennen: <Beleg-ID>_<Lieferant>[_n].<Endung>
    try {

      const ids = extractDriveFileIds(values[COL.UPLOAD - 1]);

      if (ids.length) {

        dateiname =
          ids.map(function(id, index) {
            const file = DriveApp.getFileById(id);
            const name = originalFileName(
              belegId, newSupplier, index, ids.length, file.getName());
            file.setName(name);
            return name;
          }).join(", ");

        sheet.getRange(n, COL.DATEINAME).setValue(dateiname);

      }

    } catch (error) {

      Logger.log("Original umbenennen: " + error);
      hint +=
        " Hinweis: Die Originaldatei konnte nicht umbenannt werden (" +
        (error && error.message ? error.message : error) + ").";

    }

    // Belegkopie trägt den Lieferanten im Namen
    hint += syncBelegCopiesSafe(n);

  }


  const amount = amountChanged ? newAmount : oldAmountRaw;

  return {
    message: "Die Angaben wurden gespeichert." + hint,
    amountRaw: typeof amount === "number" ? amount.toFixed(2) : String(amount || ""),
    values: {
      einreicher: text(data.einreicher),
      gekauft: text(data.gekauft),
      zweck: text(data.zweck),
      lieferant: newSupplier,
      betrag: formatAmountDisplay(amount),
      auszahlungAn: text(data.auszahlungAn),
      dateiname: dateiname
    }
  };

}



/************************************************************
 * NEUER VORGANG – EINGABESEITE
 *
 * Zweiter Eingabeweg neben dem Google-Formular, z. B. für
 * Kautionen, Eigenbelege oder Wertgutachten.
 * Der Vorgang wird direkt in CONFIG.SHEET_NAME geschrieben,
 * bekommt eine normale Beleg-ID und Status "Neu"
 * (Einnahmen: "Offen – Zahlung erwartet"). Es wird keine
 * Mail verschickt.
 ************************************************************/


function getActiveUserEmail() {

  try {
    return Session.getActiveUser().getEmail() || "";
  } catch (e) {
    return "";
  }

}


function createNewEntryPage() {

  const options =
    VORGANGSARTEN.map(function(art) {
      return (
        "<option value=\"" + htmlEscape(art) + "\"" +
        (art === VORGANGSART_STANDARD ? " selected" : "") +
        ">" + htmlEscape(art) + "</option>"
      );
    }).join("");

  return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Neuer Vorgang – KGV Neuenhof</title>

<style>

${commonCss()}

.payee-wrapper {
  position: relative;
}

.payee-suggestions {
  position: absolute;
  z-index: 50;
  left: 0;
  right: 0;
  background: white;
  border: 1px solid #ccc;
  border-radius: 0 0 7px 7px;
  max-height: 220px;
  overflow-y: auto;
  display: none;
}

.payee-suggestion {
  padding: 10px;
  cursor: pointer;
  border-bottom: 1px solid #eee;
}

.payee-suggestion:hover {
  background: #f1f3f4;
}

.small {
  font-size: 13px;
  color: #666;
}

.hidden {
  display: none;
}

</style>

<script>

const MAX_FILE_SIZE = 10 * 1024 * 1024;


function el(id) {
  return document.getElementById(id);
}


function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}


// Felder je nach Vorgangsart und Einnahme ein-/ausblenden
function updateFields() {

  const art = el("vorgangsart").value;
  const einnahme = el("einnahme").checked;

  el("rowNachkontrolle").classList.toggle(
    "hidden", art !== "Auszahlung Kaution nach Nachkontrolle");

  el("rowBegruendung").classList.toggle(
    "hidden", art !== "Eigenbeleg");

  el("payoutSection").classList.toggle(
    "hidden", einnahme);

  el("labelAuszahlungAn").innerText =
    art === "Quittung Wertgutachten"
      ? "Gutachter (Zahlungsempfänger)"
      : "Auszahlung an (Zahlungsempfänger)";

  el("labelLieferant").innerText =
    einnahme
      ? "Zahler (von wem kommt das Geld?)"
      : "Lieferant / Aussteller";

}


function updatePayeeSuggestions() {

  const input = el("auszahlungAn");
  const box = el("payeeSuggestions");
  const query = input.value.trim();

  box.innerHTML = "";

  if (!query) {
    box.style.display = "none";
    return;
  }

  google.script.run
    .withSuccessHandler(function(matches) {

      if (input.value.trim() !== query) return;

      box.innerHTML = "";

      if (!matches || !matches.length) {
        box.style.display = "none";
        return;
      }

      matches.forEach(function(item) {

        const div = document.createElement("div");
        div.className = "payee-suggestion";
        div.innerHTML =
          "<strong>" + escapeHtml(item.name) + "</strong><br>" +
          "<span class='small'>" + escapeHtml(item.iban) +
          (item.bic ? " · " + escapeHtml(item.bic) : "") + "</span>";

        div.onclick = function() {
          el("auszahlungAn").value = item.name;
          el("iban").value = item.iban || "";
          el("bic").value = item.bic || "";
          box.style.display = "none";
        };

        box.appendChild(div);

      });

      box.style.display = "block";

    })
    .withFailureHandler(function() {
      box.style.display = "none";
    })
    .getPayeeSuggestions(query);

}


function readFile(file) {

  return new Promise(function(resolve, reject) {

    if (file.size > MAX_FILE_SIZE) {
      reject(new Error("Die Datei " + file.name + " ist größer als 10 MB."));
      return;
    }

    const reader = new FileReader();

    reader.onload = function() {
      resolve({
        name: file.name,
        mimeType: file.type || "application/octet-stream",
        data: String(reader.result).split(",")[1]
      });
    };

    reader.onerror = function() {
      reject(new Error("Die Datei " + file.name + " konnte nicht gelesen werden."));
    };

    reader.readAsDataURL(file);

  });

}


function saveEntry() {

  const button = el("saveButton");
  const files = Array.prototype.slice.call(el("files").files || []);

  button.disabled = true;
  button.innerText = "Wird gespeichert...";

  Promise.all(files.map(readFile))
    .then(function(fileData) {

      const data = {
        vorgangsart: el("vorgangsart").value,
        einnahme: el("einnahme").checked,
        einreicher: el("einreicher").value,
        parzelle: el("parzelle").value,
        gekauft: el("gekauft").value,
        zweck: el("zweck").value,
        lieferant: el("lieferant").value,
        betrag: el("betrag").value,
        auszahlungAn: el("auszahlungAn").value,
        iban: el("iban").value,
        bic: el("bic").value,
        masterUpdate: el("masterUpdate").checked,
        datumNachkontrolle: el("datumNachkontrolle").value,
        begruendung: el("begruendung").value,
        bemerkung: el("bemerkung").value,
        files: fileData
      };

      google.script.run
        .withSuccessHandler(function(result) {

          el("entryForm").classList.add("hidden");
          el("successBox").classList.remove("hidden");
          el("successText").innerHTML =
            "✅ " + escapeHtml(result.message);
          el("openLink").href = result.url;

        })
        .withFailureHandler(function(error) {

          button.disabled = false;
          button.innerText = "💾 Vorgang speichern";
          alert("Fehler: " + (error && error.message ? error.message : error));

        })
        .createManualEntry(data);

    })
    .catch(function(error) {

      button.disabled = false;
      button.innerText = "💾 Vorgang speichern";
      alert("Fehler: " + error.message);

    });

}


function newEntry() {

  ["parzelle", "gekauft", "zweck", "lieferant", "betrag", "auszahlungAn",
   "iban", "bic", "datumNachkontrolle", "begruendung", "bemerkung", "files"]
    .forEach(function(id) { el(id).value = ""; });

  el("einnahme").checked = false;
  el("masterUpdate").checked = false;
  el("saveButton").disabled = false;
  el("saveButton").innerText = "💾 Vorgang speichern";

  el("successBox").classList.add("hidden");
  el("entryForm").classList.remove("hidden");

  updateFields();
  window.scrollTo(0, 0);

}


window.addEventListener("load", updateFields);

document.addEventListener("click", function(event) {
  if (!event.target.closest(".payee-wrapper")) {
    el("payeeSuggestions").style.display = "none";
  }
});

</script>

</head>

<body>

<div class="container">

<h1>
Neuer Vorgang
</h1>

<p class="small">
Alle Felder sind freiwillig. Der Vorgang erscheint danach
mit Status „Neu“ in der Belegübersicht und wird dort wie
ein Formular-Beleg weiterbearbeitet. Es wird keine E-Mail
verschickt.
</p>


<div id="entryForm">

<div class="section">

<div class="form-row">
<div class="form-label">Vorgangsart</div>
<select id="vorgangsart" onchange="updateFields()">
${options}
</select>
</div>

<div class="checkbox-row" style="margin-top:0">
<label>
<input type="checkbox" id="einnahme" onchange="updateFields()">
Einnahme (Geld kommt in die Vereinskasse)
</label>
</div>

</div>


<div class="section">

<div class="section-title">Angaben</div>

<div class="form-row">
<div class="form-label">Einreicher / Mitglied</div>
<input id="einreicher" value="${htmlEscape(getActiveUserEmail())}">
</div>

<div class="form-row">
<div class="form-label">Parzelle</div>
<input id="parzelle" placeholder="z. B. 42">
</div>

<div class="form-row">
<div class="form-label">Was wurde gekauft? / Gegenstand</div>
<input id="gekauft">
</div>

<div class="form-row">
<div class="form-label">Verwendungszweck</div>
<input id="zweck">
</div>

<div class="form-row">
<div class="form-label" id="labelLieferant">Lieferant / Aussteller</div>
<input id="lieferant">
</div>

<div class="form-row">
<div class="form-label">Betrag (€)</div>
<input id="betrag" inputmode="decimal" placeholder="z. B. 150,00">
</div>

<div class="form-row hidden" id="rowNachkontrolle">
<div class="form-label">Datum der Nachkontrolle</div>
<input type="date" id="datumNachkontrolle">
</div>

<div class="form-row hidden" id="rowBegruendung">
<div class="form-label">Begründung: Warum liegt kein Originalbeleg vor?</div>
<textarea id="begruendung"></textarea>
</div>

<div class="form-row">
<div class="form-label">Bemerkung</div>
<textarea id="bemerkung"></textarea>
</div>

</div>


<div class="section" id="payoutSection">

<div class="section-title">Zahlungsempfänger</div>

<div class="form-row">
<div class="form-label" id="labelAuszahlungAn">Auszahlung an (Zahlungsempfänger)</div>
<div class="payee-wrapper">
<input
  id="auszahlungAn"
  autocomplete="off"
  oninput="updatePayeeSuggestions()"
  placeholder="Name eingeben"
>
<div id="payeeSuggestions" class="payee-suggestions"></div>
</div>
</div>

<div class="form-row">
<div class="form-label">IBAN</div>
<input id="iban">
</div>

<div class="form-row">
<div class="form-label">BIC</div>
<input id="bic">
</div>

<div class="checkbox-row">
<label>
<input type="checkbox" id="masterUpdate">
Zahlungsempfänger in den Stammdaten speichern/aktualisieren
</label>
</div>

</div>


<div class="section">

<div class="section-title">Beleg</div>

<input type="file" id="files" multiple accept="application/pdf,image/*">

<div class="small" style="margin-top:8px">
PDF oder Foto, max. 10 MB pro Datei. Mehrere Dateien möglich.
</div>

</div>


<button
  id="saveButton"
  class="button save-button"
  onclick="saveEntry()"
>
💾 Vorgang speichern
</button>

</div>


<div id="successBox" class="hidden">

<div class="message" id="successText"></div>

<a
  id="openLink"
  class="button document-button"
  href="#"
  target="_blank"
>
📄 Vorgang öffnen
</a>

<button
  class="button secondary-button"
  onclick="newEntry()"
>
+ Weiteren Vorgang erfassen
</button>

</div>


<a
  class="button back-button"
  href="${CONFIG.WEBAPP_URL}?uebersicht=1"
  target="_blank"
>
← Zur Belegübersicht
</a>

</div>

</body>

</html>

`;

}


function safeFileNamePart(value, fallback) {

  return (
    String(value || "")
      .trim()
      .replace(/[^\wäöüÄÖÜß-]+/g, "_")
      .replace(/^_+|_+$/g, "")
      .substring(0, 60) ||
    fallback
  );

}


// Name der Originaldatei: <Beleg-ID>_<Lieferant>[_n].<Endung>
function originalFileName(belegId, supplier, index, count, currentName) {

  return (
    belegId + "_" +
    safeFileNamePart(supplier, "Beleg") +
    (count > 1 ? "_" + (index + 1) : "") +
    getFileExtension(currentName)
  );

}


function createManualEntry(data) {

  data = data || {};

  const sheet = getSheet();

  const art =
    VORGANGSARTEN.indexOf(data.vorgangsart) >= 0
      ? data.vorgangsart
      : VORGANGSART_STANDARD;

  const einnahme = !!data.einnahme;
  const betrag = parseAmount(data.betrag);
  const text = function(value) { return String(value || "").trim(); };


  const zusatz = [];

  if (text(data.parzelle)) {
    zusatz.push("Parzelle: " + text(data.parzelle));
  }

  if (art === "Auszahlung Kaution nach Nachkontrolle" && text(data.datumNachkontrolle)) {
    zusatz.push(
      "Nachkontrolle am: " +
      formatDate(parseDateInput(text(data.datumNachkontrolle)))
    );
  }

  if (art === "Eigenbeleg" && text(data.begruendung)) {
    zusatz.push("Begründung Eigenbeleg: " + text(data.begruendung));
  }

  if (text(data.bemerkung)) {
    zusatz.push("Bemerkung: " + text(data.bemerkung));
  }


  // Dateien zuerst speichern: schlägt der Upload fehl,
  // wird kein halber Vorgang angelegt.
  const files = data.files || [];

  const uploadFolder =
    files.length ? getBelegFolder(CONFIG.BELEG_ORDNER_ORIGINALE) : null;

  const uploadedFiles =
    files.map(function(file) {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(file.data),
        file.mimeType,
        file.name
      );
      return uploadFolder.createFile(blob);
    });


  const now = new Date();
  const lock = LockService.getScriptLock();
  let belegId;
  let row;

  lock.waitLock(30000);

  try {

    belegId = nextBelegId(sheet, now.getFullYear());

    const values = new Array(COL_MAX).fill("");

    values[COL.ZEITSTEMPEL - 1] = now;
    values[COL.EINREICHER - 1] = text(data.einreicher);
    values[COL.GEKAUFT - 1] = text(data.gekauft);
    values[COL.ZWECK - 1] = text(data.zweck);
    values[COL.LIEFERANT - 1] = text(data.lieferant);
    values[COL.BETRAG - 1] = betrag;
    values[COL.EMAIL - 1] = getActiveUserEmail();
    values[COL.BELEG_ID - 1] = belegId;
    values[COL.STATUS - 1] = einnahme ? STATUS_EINNAHME_OFFEN : "Neu";
    values[COL.LETZTE_BEARBEITUNG - 1] = "Über Eingabeseite erfasst";
    values[COL.INTERNE_ID - 1] = Utilities.getUuid();
    values[COL.VORGANGSART - 1] = art;
    values[COL.ZUSATZANGABEN - 1] = zusatz.join("\n");
    values[COL.EINNAHME - 1] = einnahme ? "Ja" : "";

    if (!einnahme) {
      values[COL.AUSZAHLUNG_AN - 1] = text(data.auszahlungAn);
      values[COL.ZAHLUNGSEMPFAENGER - 1] = text(data.auszahlungAn);
      values[COL.IBAN - 1] = text(data.iban);
      values[COL.BIC - 1] = text(data.bic);
      values[COL.ZAHLUNGSZWECK - 1] = text(data.zweck);
    }

    if (uploadedFiles.length) {

      values[COL.UPLOAD - 1] =
        uploadedFiles.map(function(file) {
          return "https://drive.google.com/open?id=" + file.getId();
        }).join(", ");

      values[COL.DATEINAME - 1] =
        uploadedFiles.map(function(file, index) {
          const name = originalFileName(
            belegId, data.lieferant, index, uploadedFiles.length, file.getName());
          file.setName(name);
          return name;
        }).join(", ");

    }

    row = sheet.getLastRow() + 1;

    sheet
      .getRange(row, 1, 1, COL_MAX)
      .setValues([values]);

  } finally {

    lock.releaseLock();

  }


  if (!einnahme && data.masterUpdate && text(data.auszahlungAn)) {
    savePayeeToMasterData(
      text(data.auszahlungAn),
      text(data.iban),
      text(data.bic)
    );
  }


  return {
    belegId: belegId,
    url: makeWebUrl(belegId),
    message:
      "Vorgang " + belegId + " (" + art + (einnahme ? ", Einnahme" : "") +
      ") wurde angelegt." +
      syncBelegCopiesSafe(row)
  };

}


/************************************************************
 * BELEGÜBERSICHT
 ************************************************************/


function createOverviewPage() {

  return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Belegübersicht KGV Neuenhof</title>

<style>

${commonCss()}

.search {
  margin-bottom: 20px;
}

.overview-table {
  overflow-x: auto;
}

.overview-table table {
  min-width: 850px;
}

.open-button {
  display: inline-block;
  padding: 8px 12px;
  background: #1769aa;
  color: white;
  border-radius: 6px;
  text-decoration: none;
  font-weight: bold;
}

select {
  margin-top: 0;
}

</style>

<script>

let rows = [];
let currentSort = "date";
let currentDirection = "desc";


function loadOverview() {

  google.script.run

    .withSuccessHandler(
      function(data) {

        rows = data;

        fillStatusFilter();

        renderTable();

      }
    )

    .withFailureHandler(
      function(error) {

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .getOverviewData();

}


function renderTable() {

  const search =
    document.getElementById(
      "search"
    ).value
      .toLowerCase();

  const status =
    document.getElementById(
      "statusFilter"
    ).value;

  const art =
    document.getElementById(
      "artFilter"
    ).value;


  let filtered =
    rows.filter(
      function(item) {

        const text =
          JSON.stringify(item)
            .toLowerCase();

        const matchesSearch =
          !search ||
          text.indexOf(search) >= 0;

        const matchesStatus =
          !status ||
          item.status === status;

        const matchesArt =
          !art ||
          item.vorgangsart === art;

        return (
          matchesSearch &&
          matchesStatus &&
          matchesArt
        );

      }
    );


  filtered.sort(
    function(a,b) {

      let av;
      let bv;


      if (
        currentSort === "id"
      ) {

        av = a.belegId;
        bv = b.belegId;

      } else if (
        currentSort === "supplier"
      ) {

        av = a.lieferant;
        bv = b.lieferant;

      } else if (
        currentSort === "person"
      ) {

        av = a.einreicher;
        bv = b.einreicher;

      } else if (
        currentSort === "amount"
      ) {

        av = Number(a.betrag) || 0;
        bv = Number(b.betrag) || 0;

      } else if (
        currentSort === "art"
      ) {

        av = a.vorgangsart;
        bv = b.vorgangsart;

      } else if (
        currentSort === "status"
      ) {

        av = a.status;
        bv = b.status;

      } else {

        av = a.timestampValue;
        bv = b.timestampValue;

      }


      if (
        av < bv
      ) {
        return currentDirection === "asc"
          ? -1
          : 1;
      }

      if (
        av > bv
      ) {
        return currentDirection === "asc"
          ? 1
          : -1;
      }

      return 0;

    }
  );


  const tbody =
    document.getElementById(
      "tbody"
    );

  tbody.innerHTML =
    "";


  filtered.forEach(
    function(item) {

      const tr =
        document.createElement(
          "tr"
        );


      tr.innerHTML =

        "<td>" +
        escapeHtml(item.belegId) +
        "</td>" +

        "<td>" +
        escapeHtml(item.timestamp) +
        "</td>" +

        "<td>" +
        escapeHtml(item.vorgangsart) +
        "</td>" +

        "<td>" +
        escapeHtml(item.einreicher) +
        "</td>" +

        "<td>" +
        escapeHtml(item.gekauft) +
        "</td>" +

        "<td>" +
        escapeHtml(item.lieferant) +
        "</td>" +

        "<td>" +
        escapeHtml(item.betrag) +
        " €</td>" +

        "<td>" +
        escapeHtml(item.status) +
        "</td>" +

        "<td>" +

        "<a " +
        "class='open-button' " +
        "href='" +
        "${CONFIG.WEBAPP_URL}?key=" +
        encodeURIComponent(
          item.internalId
        ) +
        "' " +
        "target='_blank'>" +

        "📄 Öffnen" +

        "</a>" +

        "</td>";


      tbody.appendChild(
        tr
      );

    }
  );


  document.getElementById(
    "count"
  ).innerText =
    filtered.length +
    " Belege";

}


function fillStatusFilter() {

  fillFilter("statusFilter", "status");
  fillFilter("artFilter", "vorgangsart");

}


function fillFilter(selectId, field) {

  const select =
    document.getElementById(
      selectId
    );

  const seen = {};

  rows.forEach(
    function(item) {

      const value = item[field];

      if (!value || seen[value]) {
        return;
      }

      seen[value] = true;

      const option =
        document.createElement(
          "option"
        );

      option.value = value;
      option.textContent = value;

      select.appendChild(
        option
      );

    }
  );

}


function sortBy(
  field
) {

  if (
    currentSort === field
  ) {

    currentDirection =
      currentDirection === "asc"
        ? "desc"
        : "asc";

  } else {

    currentSort =
      field;

    currentDirection =
      "asc";

  }

  renderTable();

}


function escapeHtml(
  value
) {

  return String(value || "")
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;")
    .replace(/'/g,"&#039;");

}


window.onload =
  loadOverview;

</script>

</head>

<body>

<div class="container">

<h1>
Belegübersicht KGV Neuenhof
</h1>


<a
  class="button save-button"
  style="margin-bottom:20px"
  href="${CONFIG.WEBAPP_URL}?neu=1"
  target="_blank"
>
+ Neuer Vorgang
</a>


<div class="search">

<input
  id="search"
  placeholder="Suche nach Beleg-ID, Datum, Einreicher, Einkauf, Zweck, Lieferant, Betrag oder Status"
  oninput="renderTable()"
/>

</div>


<div class="search">

<select
  id="statusFilter"
  onchange="renderTable()"
>

<option value="">
Alle Status
</option>

</select>

</div>


<div class="search">

<select
  id="artFilter"
  onchange="renderTable()"
>

<option value="">
Alle Vorgangsarten
</option>

</select>

</div>


<div id="count"
style="margin-bottom:15px">
</div>


<div class="overview-table">

<table>

<thead>

<tr>

<th
onclick="sortBy('id')"
style="cursor:pointer"
>
Beleg-ID
</th>

<th
onclick="sortBy('date')"
style="cursor:pointer"
>
Datum
</th>

<th
onclick="sortBy('art')"
style="cursor:pointer"
>
Vorgang
</th>

<th
onclick="sortBy('person')"
style="cursor:pointer"
>
Einreicher
</th>

<th>
Gekauft
</th>

<th
onclick="sortBy('supplier')"
style="cursor:pointer"
>
Lieferant
</th>

<th
onclick="sortBy('amount')"
style="cursor:pointer"
>
Betrag
</th>

<th
onclick="sortBy('status')"
style="cursor:pointer"
>
Status
</th>

<th>
Aktion
</th>

</tr>

</thead>

<tbody id="tbody">
</tbody>

</table>


</div>




</div>

</body>

</html>

`;

}


/************************************************************
 * DATEN FÜR BELEGÜBERSICHT
 ************************************************************/


function getOverviewData() {
  ensureAllInternalIds();

  const sheet =
    getSheet();

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        COL_MAX
      )
      .getValues();


  return values
    .filter(
      function(row) {

        return (
          String(
            row[
              COL.BELEG_ID - 1
            ]
          ).trim() !== ""
        );

      }
    )
    .map(
      function(row) {

        return {

          belegId:
            String(
              row[
                COL.BELEG_ID - 1
              ] || ""
            ),

          internalId:
            String(
              row[
                COL.INTERNE_ID - 1
              ] || ""
            ),

          timestamp:
            formatDateTime(
              row[
                COL.ZEITSTEMPEL - 1
              ]
            ),

          timestampValue:
            row[
              COL.ZEITSTEMPEL - 1
            ]
              ? new Date(
                  row[
                    COL.ZEITSTEMPEL - 1
                  ]
                ).getTime()
              : 0,

          einreicher:
            String(
              row[
                COL.EINREICHER - 1
              ] || ""
            ),

          gekauft:
            String(
              row[
                COL.GEKAUFT - 1
              ] || ""
            ),

          zweck:
            String(
              row[
                COL.ZWECK - 1
              ] || ""
            ),

          lieferant:
            String(
              row[
                COL.LIEFERANT - 1
              ] || ""
            ),

          betrag:
            row[
              COL.BETRAG - 1
            ],

          status:
            String(
              row[
                COL.STATUS - 1
              ] || ""
            ),

          vorgangsart:
            getVorgangsart(row) +
            (isEinnahme(row) ? " (Einnahme)" : "")

        };

      }
    );

}


/************************************************************
 * VORSITZENDEN-ÜBERSICHT
 ************************************************************/


function createChairmanOverviewPage() {

  return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Meine Genehmigungen</title>

<style>

${commonCss()}

</style>

<script>

function load() {

  google.script.run

    .withSuccessHandler(
      function(data) {

        const body =
          document.getElementById(
            "body"
          );

        body.innerHTML =
          "";

        data.forEach(
          function(item) {

            const div =
              document.createElement(
                "div"
              );

            div.className =
              "section";

            div.innerHTML =

              "<strong>" +
              escapeHtml(
                item.belegId
              ) +
              "</strong><br>" +

              "Lieferant: " +
              escapeHtml(
                item.lieferant
              ) +
              "<br>" +

              "Betrag: " +
              escapeHtml(
                item.betrag
              ) +
              " €<br>" +

              "Status: " +
              escapeHtml(
                item.status
              ) +

              "<a " +
              "class='button document-button' " +
              "href='" +
              "${CONFIG.WEBAPP_URL}?id=" +
              encodeURIComponent(
                item.belegId
              ) +
              "&rolle=vorsitzender' " +
              "target='_blank'>" +
              "📄 Öffnen" +
              "</a>";

            body.appendChild(
              div
            );

          }
        );

        if (!data.length) {

          body.innerHTML =
            "<div class='message'>" +
            "Keine Genehmigungen vorhanden." +
            "</div>";

        }

      }
    )

    .withFailureHandler(
      function(error) {

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .getChairmanOverviewData();

}


function escapeHtml(
  value
) {

  return String(value || "")
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;")
    .replace(/'/g,"&#039;");

}


window.onload =
  load;

</script>

</head>

<body>

<div class="container">

<h1>
Meine Genehmigungen
</h1>

<div id="body">
Wird geladen...
</div>


<a
  class="button back-button"
  href="${CONFIG.WEBAPP_URL}?uebersicht=1"
  target="_blank"
>
← Belegübersicht
</a>

</div>

</body>

</html>

`;

}


function getChairmanOverviewData() {

  const sheet =
    getSheet();

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        COL.GEBUCHT_AM
      )
      .getValues();


  return values
    .filter(
      function(row) {

        return (
          String(
            row[
              COL.STATUS - 1
            ]
          ).indexOf(
            "Genehmigung Vorsitzender"
          ) >= 0 ||

          (
            row[
              COL.GENEHMIGUNG_VORSITZENDER - 1
            ]
          )

        );

      }
    )
    .map(
      function(row) {

        return {

          belegId:
            row[
              COL.BELEG_ID - 1
            ],

          lieferant:
            row[
              COL.LIEFERANT - 1
            ],

          betrag:
            row[
              COL.BETRAG - 1
            ],

          status:
            row[
              COL.STATUS - 1
            ]

        };

      }
    );

}


/************************************************************
 * 2. KASSIERER – EIGENE ÜBERSICHT
 ************************************************************/


function createPaymentOverview() {

  return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Meine Überweisungen</title>

<style>

${commonCss()}

</style>

<script>

function load() {

  google.script.run

    .withSuccessHandler(
      function(data) {

        const body =
          document.getElementById(
            "body"
          );

        body.innerHTML =
          "";

        data.forEach(
          function(item) {

            const div =
              document.createElement(
                "div"
              );

            div.className =
              "section";

            div.innerHTML =

              "<strong>" +
              escapeHtml(
                item.belegId
              ) +
              "</strong><br>" +

              "Lieferant: " +
              escapeHtml(
                item.lieferant
              ) +
              "<br>" +

              "Betrag: " +
              escapeHtml(
                item.betrag
              ) +
              " €<br>" +

              "Status: " +
              escapeHtml(
                item.status
              ) +

              "<a " +
              "class='button document-button' " +
              "href='" +
              "${CONFIG.WEBAPP_URL}?id=" +
              encodeURIComponent(
                item.belegId
              ) +
              "&rolle=zweiter_kassierer' " +
              "target='_blank'>" +
              "📄 Öffnen" +
              "</a>";

            body.appendChild(
              div
            );

          }
        );


        if (!data.length) {

          body.innerHTML =
            "<div class='message'>" +
            "Keine Überweisungen vorhanden." +
            "</div>";

        }

      }
    )

    .withFailureHandler(
      function(error) {

        alert(
          "Fehler: " +
          error.message
        );

      }
    )

    .getPaymentOverviewData();

}


function escapeHtml(
  value
) {

  return String(value || "")
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;")
    .replace(/'/g,"&#039;");

}


window.onload =
  load;

</script>

</head>

<body>

<div class="container">

<h1>
Meine Überweisungen
</h1>

<div id="body">
Wird geladen...
</div>


<a
  class="button back-button"
  href="${CONFIG.WEBAPP_URL}?uebersicht=1"
  target="_blank"
>
← Belegübersicht
</a>

</div>

</body>

</html>

`;

}


/************************************************************
 * DATEN 2. KASSIERER
 ************************************************************/


function getPaymentOverviewData() {

  const sheet =
    getSheet();

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }


  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        COL.GEBUCHT_AM
      )
      .getValues();


  return values
    .filter(
      function(row) {

        const sent =
          row[
            COL.AN_2_KASSIERER_GESENDET - 1
          ];

        return !!sent;

      }
    )
    .map(
      function(row) {

        return {

          belegId:
            row[
              COL.BELEG_ID - 1
            ],

          lieferant:
            row[
              COL.LIEFERANT - 1
            ],

          betrag:
            row[
              COL.BETRAG - 1
            ],

          status:
            row[
              COL.STATUS - 1
            ]

        };

      }
    );

}


/************************************************************
 * FEHLERSEITE
 ************************************************************/


function errorPage(
  message
) {

  return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<style>

body {
  font-family: Arial, sans-serif;
  background: #f5f7fa;
  padding: 20px;
}

.container {
  max-width: 700px;
  margin: auto;
  background: white;
  padding: 30px;
  border-radius: 12px;
  box-shadow: 0 2px 10px rgba(0,0,0,.12);
}

</style>

</head>

<body>

<div class="container">

<h1>
Belegprüfung KGV Neuenhof
</h1>

<p>
${htmlEscape(message)}
</p>

</div>

</body>

</html>

`;

}

/************************************************************
 * DIAGNOSE
 ************************************************************/
function diagnoseBeleg(identifier) {
  const byId=getRowData(identifier,false);
  const byKey=getRowData(identifier,true);
  return {identifier:String(identifier||""),byId:byId?{row:byId.rowNumber,belegId:byId.values[COL.BELEG_ID-1],internalId:byId.values[COL.INTERNE_ID-1],timestamp:String(byId.values[COL.ZEITSTEMPEL-1])}:null,byKey:byKey?{row:byKey.rowNumber,belegId:byKey.values[COL.BELEG_ID-1],internalId:byKey.values[COL.INTERNE_ID-1],timestamp:String(byKey.values[COL.ZEITSTEMPEL-1])}:null};
}

