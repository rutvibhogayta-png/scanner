//==================================================
// CONFIGURATION
//==================================================

const SHEET_GUESTS = "Guests";
const SHEET_LOGS = "MealLogs";
const SHEET_SLOTS = "MealSlots";


//==================================================
// GET REQUESTS
//==================================================

function doGet(e) {

  const action = e && e.parameter
    ? e.parameter.action
    : "";


  //==============================================
  // MEAL SLOTS API
  //==============================================

  if (action === "slots") {

    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_SLOTS);

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {

      return json({
        status: "ok",
        slots: []
      });

    }

    const data = sheet
      .getRange(2, 1, lastRow - 1, 2)
      .getValues();

    const slots = data.map(row => ({
      key: row[0],
      label: row[1]
    }));

    return json({
      status: "ok",
      slots: slots
    });
  }


  //==============================================
  // OPEN SCANNER PAGE
  //==============================================

  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("Meal Check-In");

}


//==================================================
// POST REQUESTS
//==================================================

function doPost(e) {

  try {

    if (!e || !e.postData) {

      return json({

        status: "error",

        message: "No POST data received."

      });

    }


    const data = JSON.parse(
      e.postData.contents
    );


    //============================================
    // SCAN
    //============================================

    if (data.action === "scan") {

      return processScan(data);

    }


    return json({

      status: "error",

      message: "Invalid action."

    });


  } catch (error) {

    return json({

      status: "error",

      message: error.message

    });

  }

}


//==================================================
// PROCESS QR SCAN
//==================================================

function processScan(data) {


  //----------------------------------------------
  // GET DATA FROM SCANNER
  //----------------------------------------------

  const guestId = String(
    data.guestId || ""
  )
    .trim()
    .toUpperCase();


  const meal = String(
    data.meal || "Meal"
  ).trim();


  const station = String(
    data.station || ""
  ).trim();


  //----------------------------------------------
  // VALIDATION
  //----------------------------------------------

  if (!guestId) {

    return json({

      status: "error",

      message: "Guest ID missing."

    });

  }


  //----------------------------------------------
  // OPEN SPREADSHEET
  //----------------------------------------------

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();


  const guestSheet =
    spreadsheet.getSheetByName(
      SHEET_GUESTS
    );


  const logSheet =
    spreadsheet.getSheetByName(
      SHEET_LOGS
    );


  if (!guestSheet) {

    return json({

      status: "error",

      message: "Guests sheet not found."

    });

  }


  if (!logSheet) {

    return json({

      status: "error",

      message: "MealLogs sheet not found."

    });

  }


  //----------------------------------------------
  // GET GUEST DATA
  //----------------------------------------------

  const lastRow =
    guestSheet.getLastRow();


  if (lastRow < 2) {

    return json({

      status: "invalid",

      message: "No guests found."

    });

  }


  /*
    Guests sheet:

    Column A = Guest ID
    Column B = Name
    Column C = Meal
    Column D = QR
  */


  const guestData =
    guestSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        4
      )
      .getValues();


  //----------------------------------------------
  // FIND GUEST
  //----------------------------------------------

  let guestRow = -1;

  let guest = null;


  for (
    let i = 0;
    i < guestData.length;
    i++
  ) {


    const currentId =
      String(guestData[i][0])
        .trim()
        .toUpperCase();


    if (currentId === guestId) {

      guestRow = i + 2;

      guest = guestData[i];

      break;

    }

  }


  //----------------------------------------------
  // GUEST NOT FOUND
  //----------------------------------------------

  if (!guest) {

    return json({

      status: "invalid",

      message: "Guest not found."

    });

  }


  //----------------------------------------------
  // GET NAME
  //----------------------------------------------

  const name = guest[1];


  //----------------------------------------------
  // CHECK IF MEAL ALREADY TAKEN
  //----------------------------------------------

  const alreadyTaken =
    guest[2] === true;


  if (alreadyTaken) {

    return json({

      status: "duplicate",

      guestId: guestId,

      name: name,

      meal: meal,

      message: "Meal already collected."

    });

  }


  //----------------------------------------------
  // MARK MEAL AS COLLECTED
  //----------------------------------------------

  guestSheet
    .getRange(guestRow, 3)
    .setValue(true);


  //----------------------------------------------
  // ADD LOG
  //----------------------------------------------

  logSheet.appendRow([

    new Date(),

    guestId,

    name,

    meal,

    station

  ]);


  //----------------------------------------------
  // SUCCESS RESPONSE
  //----------------------------------------------

  return json({

    status: "success",

    guestId: guestId,

    name: name,

    meal: meal,

    message: "Meal successfully collected."

  });

}


//==================================================
// GENERATE LARGE QR CODES
//==================================================

function generateQRCodes() {


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_GUESTS
      );


  if (!sheet) {

    throw new Error(
      "Guests sheet not found."
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {

    return;

  }


  //----------------------------------------------
  // GET GUEST IDS
  //----------------------------------------------

  const guestIds =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        1
      )
      .getValues();


  //----------------------------------------------
  // QR COLUMN WIDTH
  //----------------------------------------------

  sheet.setColumnWidth(
    4,
    220
  );


  //----------------------------------------------
  // GENERATE EACH QR
  //----------------------------------------------

  for (
    let i = 0;
    i < guestIds.length;
    i++
  ) {


    const guestId =
      String(guestIds[i][0])
        .trim()
        .toUpperCase();


    if (!guestId) {

      continue;

    }


    //--------------------------------------------
    // QUICKCHART QR
    //--------------------------------------------

    const qrUrl =
      "https://quickchart.io/qr" +
      "?size=800" +
      "&margin=2" +
      "&text=" +
      encodeURIComponent(
        guestId
      );


    //--------------------------------------------
    // INSERT QR
    //--------------------------------------------

    sheet
      .getRange(
        i + 2,
        4
      )
      .setFormula(
        `=IMAGE("${qrUrl}",4,200,200)`
      );


    //--------------------------------------------
    // ROW HEIGHT
    //--------------------------------------------

    sheet.setRowHeight(
      i + 2,
      220
    );

  }

}


//==================================================
// JSON RESPONSE
//==================================================

function json(obj) {

  return ContentService

    .createTextOutput(
      JSON.stringify(obj)
    )

    .setMimeType(
      ContentService.MimeType.JSON
    );

}
