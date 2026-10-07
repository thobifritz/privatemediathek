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
  INTERNE_ID: 34

};



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

  return sheet;
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
  const values=sheet.getRange(rowNumber,1,1,Math.max(sheet.getLastColumn(),COL.INTERNE_ID)).getValues()[0];
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
    s.indexOf("neu") >= 0
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

      const id =
        String(item[0] || "").trim();

      const match =
        id.match(new RegExp("^(\\d{4})-(\\d+)$"));

      if (!match) {
        return;
      }

      if (Number(match[1]) !== Number(year)) {
        return;
      }

      const number =
        Number(match[2]);

      if (number > highestNumber) {
        highestNumber = number;
      }

    });

    const nextNumber =
      highestNumber + 1;

    belegId =
      year +
      "-" +
      Utilities.formatString(
        "%04d",
        nextNumber
      );

    sheet
      .getRange(row, COL.BELEG_ID)
      .setValue(belegId);

    sheet
      .getRange(row, COL.STATUS)
      .setValue("Neu");

  } finally {

    lock.releaseLock();

  }


  sheet
    .getRange(
      row,
      COL.BELEG_ID
    )
    .setValue(
      belegId
    );


  sheet
    .getRange(
      row,
      COL.STATUS
    )
    .setValue(
      "Neu"
    );

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

<title>Belegprüfung KGV Neuenhof – DIAGNOSE 2026-09-29</title>

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
  ${jsonForHtml(String(belegId))};

const rowNumber =
  ${jsonForHtml(result.rowNumber)};

const sepaAmount =
  ${jsonForHtml(String(betrag || "").replace(",", "."))};

const payees =
  ${jsonForHtml(payeeData)};



function buildSepaPayload() {
  const recipient = document.getElementById("zahlungsempfaenger");
  const ibanEl = document.getElementById("iban");
  const bicEl = document.getElementById("bic");
  const purposeEl = document.getElementById("zahlungszweck");

  if (!recipient || !ibanEl || !purposeEl) return "";

  const name = recipient.value.trim();
  const iban = ibanEl.value.replace(/\s+/g, "").trim();
  const bic = bicEl ? bicEl.value.replace(/\s+/g, "").trim() : "";
  const purpose = purposeEl.value.trim();

  if (!name || !iban || !purpose) return "";

  const amount = String(sepaAmount || "").replace(",", ".").trim();

  return [
    "BCD",
    "002",
    "1",
    "SCT",
    bic,
    name,
    iban,
    amount ? "EUR" + amount : "EUR0.00",
    purpose
  ].join("\n");
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


function diagnoseButton() {
  const button = document.getElementById("diagnoseButton");
  const result = document.getElementById("diagnoseResult");

  if (button) {
    button.disabled = true;
    button.innerText = "Teste Verbindung...";
  }

  if (result) {
    result.style.display = "block";
    result.innerHTML = "⏳ Test läuft...";
  }

  if (typeof google === "undefined" || !google.script || !google.script.run) {
    if (result) {
      result.innerHTML =
        "❌ google.script.run ist im Browser nicht verfügbar.";
    }
    if (button) {
      button.disabled = false;
      button.innerText = "🔧 Verbindung testen";
    }
    return;
  }

  google.script.run
    .withSuccessHandler(function(data) {
      if (result) {
        if (data && data.ok) {
          result.innerHTML =
            "✅ " + data.message +
            "<br>Schlüssel: " + data.identifier +
            "<br>Zeile: " + data.row +
            "<br>Beleg-ID: " + data.belegId;
        } else {
          result.innerHTML =
            "⚠️ " + ((data && data.message) || "Unbekannte Antwort.");
        }
      }

      if (button) {
        button.disabled = false;
        button.innerText = "🔧 Verbindung erneut testen";
      }
    })
    .withFailureHandler(function(error) {
      if (result) {
        result.innerHTML =
          "❌ Serverfehler: " +
          (error && error.message ? error.message : String(error));
      }

      if (button) {
        button.disabled = false;
        button.innerText = "🔧 Verbindung erneut testen";
      }
    })
    .diagnoseConnection(belegId);
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

    .originalbelegVorhanden(
      belegId
    );

}


function savePaymentData() {

  const data = {

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
      rowNumber
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


["zahlungsempfaenger","iban","bic","zahlungszweck"].forEach(function(id) {
  const field = document.getElementById(id);
  if (field) {
    field.addEventListener("input", renderSepaQrCode);
  }
});

window.addEventListener("load", renderSepaQrCode);

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


function chooseAction(
  action
) {

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


document.addEventListener("DOMContentLoaded", function() {
  const btn = document.getElementById("diagnoseButton");
  if (btn) {
    btn.addEventListener("click", diagnoseButton);
  }
});

</script>

</head>

<body>

<div class="container">

<h1>
Belegprüfung KGV Neuenhof
</h1>

<div style="
  margin:15px 0;
  padding:15px;
  border:2px solid #d9a400;
  border-radius:10px;
  background:#fff8db;
">
  <strong>Diagnoseversion 3 – 2026-09-29</strong><br>
  <span style="font-size:13px;">
    Dieser erste Test benutzt weder Apps Script noch google.script.run.
  </span>

  <button
    id="simpleTestButton"
    class="button"
    style="background:#d9a400;color:#111;"
    onclick="
      this.innerText='✅ JAVASCRIPT-KLICK FUNKTIONIERT';
      this.style.background='#188038';
      document.getElementById('simpleTestResult').innerText='Der Browser verarbeitet den Klick.';
    "
  >
    🧪 Einfachen Klick testen
  </button>

  <div
    id="simpleTestResult"
    style="
      margin-top:10px;
      padding:10px;
      background:white;
      border-radius:6px;
    "
  >
    Noch nicht getestet.
  </div>

  <button
    id="diagnoseButton"
    class="button"
    style="background:#666;color:white;"
    onclick="
      this.disabled=true;
      this.innerText='⏳ Server wird getestet...';
      google.script.run
        .withSuccessHandler(function(data){
          document.getElementById('diagnoseResult').style.display='block';
          document.getElementById('diagnoseResult').innerHTML =
            data && data.ok
              ? '✅ SERVER ERREICHBAR – Beleg-ID: ' + data.belegId + ' – Zeile: ' + data.row
              : '⚠️ SERVER ERREICHBAR, ABER: ' + ((data && data.message) || 'unbekannte Antwort');
          this.disabled=false;
          this.innerText='🔧 Verbindung erneut testen';
        }.bind(this))
        .withFailureHandler(function(error){
          document.getElementById('diagnoseResult').style.display='block';
          document.getElementById('diagnoseResult').innerHTML =
            '❌ SERVERFEHLER: ' + (error && error.message ? error.message : String(error));
          this.disabled=false;
          this.innerText='🔧 Verbindung erneut testen';
        }.bind(this))
        .diagnoseConnection(belegId);
    "
  >
    🔧 Danach Verbindung testen
  </button>

  <div
    id="diagnoseResult"
    style="
      display:none;
      margin-top:10px;
      padding:10px;
      background:white;
      border-radius:6px;
    "
  ></div>
</div>


<div class="beleg-id">
Beleg ${htmlEscape(belegId)}
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
<td class="label">Rechnungsnummer</td>
<td>${htmlEscape(rechnungsnummer)}</td>
</tr>

<tr>
<td class="label">E-Mail</td>
<td>${htmlEscape(email)}</td>
</tr>

<tr>
<td class="label">Dateiname</td>
<td>${htmlEscape(dateiname)}</td>
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


<div class="section">

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


<button
  class="button decision-button"
  onclick="chooseAction('Genehmigung Vorsitzender')"
>
→ An 1. Vorsitzenden zur Genehmigung
</button>


<button
  class="button decision-button"
  onclick="chooseAction('Zahlung an 2. Kassierer')"
>
→ Direkt an 2. Kassierer zur Überweisung
</button>


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
Bezahlt am
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
Bezahlt am
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
✓ Als bezahlt markieren
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

function savePaymentDataByRow(rowNumber, data) {
  return savePaymentData(getBelegIdFromRow(rowNumber), data || {});
}

function sendToSecondCashierByRow(rowNumber) {
  return sendToSecondCashier(getBelegIdFromRow(rowNumber));
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
    (
      rechnungsnummer
        ? rechnungsnummer
        : belegId
    );


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

    "Rechnungsnummer: " +
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

    "Rechnungsnummer: " +
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
  ${jsonForHtml(String(belegId))};

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

        document.getElementById(
          "result"
        ).innerHTML =
          "✅ " +
          result;

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

</script>

</head>

<body>

<div class="container">

<h1>
Beleggenehmigung KGV Neuenhof
</h1>

<div class="beleg-id">
Beleg ${htmlEscape(belegId)}
</div>


<div
  class="status-box ${getStatusClass(status)}"
>
Status:
${htmlEscape(status)}
</div>


<table style="margin-top:20px">

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
<td class="label">Rechnungsnummer</td>
<td>${htmlEscape(rechnungsnummer)}</td>
</tr>

<tr>
<td class="label">E-Mail</td>
<td>${htmlEscape(email)}</td>
</tr>

<tr>
<td class="label">Dateiname</td>
<td>${htmlEscape(filename)}</td>
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
    "BCD\n002\n1\nSCT\n" +
    String(bic || "") + "\n" +
    String(recipient || "") + "\n" +
    String(iban || "").replace(/\s+/g, "") + "\n" +
    "EUR" + String(betrag || "").replace(",", ".") + "\n" +
    String(paymentPurpose || "")
  )};

function renderPaymentQr() {
  const box = document.getElementById("sepaQrCodePayment");
  if (!box) return;

  const payload = String(sepaPaymentPayload || "");

  if (!payload) {
    box.innerHTML =
      "<div class='small'>Keine vollständigen SEPA-Daten vorhanden.</div>";
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
  ${jsonForHtml(String(belegId))};

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
Beleg ${htmlEscape(belegId)}
</div>


<div
  class="status-box ${getStatusClass(status)}"
>
Status:
${htmlEscape(status)}
</div>


<table style="margin-top:20px">

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
<td class="label">Rechnungsnummer</td>
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
Zahlungsdaten durchführen.
</p>

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


  sheet
    .getRange(
      row,
      COL.STATUS
    )
    .setValue(
      "Bezahlt"
    );


  sheet
    .getRange(
      row,
      COL.LETZTE_BEARBEITUNG
    )
    .setValue(
      "Als bezahlt markiert"
    );


  return (
    "Der Beleg wurde als bezahlt markiert. " +
    "Zahlungsdatum: " +
    Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "dd.MM.yyyy"
    )
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
    )
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
    "Die Datumsangaben wurden gespeichert."
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


  sheet.deleteRow(
    row
  );


  return (
    "Beleg " +
    belegId +
    " wurde aus der Belegliste gelöscht."
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

        return (
          matchesSearch &&
          matchesStatus
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
<div id="sepaQrCodePayment" style="text-align:center;padding:10px">
  <div class="small">SEPA-QR-Code</div>
</div>


</div>


<a
  class="button back-button"
  href="${CONFIG.WEBAPP_URL}"
>
Startseite
</a>


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
        COL.INTERNE_ID
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
            )

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

