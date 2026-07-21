//=========================
// CONFIGURATION
//=========================

const AUTH_CODE = "MEAL2026";

const SHEET_GUESTS = "Guests";
const SHEET_LOGS = "MealLogs";
const SHEET_SLOTS = "MealSlots";


//=========================
// GET REQUESTS
//=========================

function doGet(e) {

  const action = e.parameter.action;

  // Return Meal Slots
  if (action == "slots") {

    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SLOTS);

    const data = sheet.getRange(2,1,sheet.getLastRow()-1,2).getValues();

    const slots = data.map(r=>({
      key:r[0],
      label:r[1]
    }));

    return json({
      status:"ok",
      slots:slots
    });
  }

  return json({
    status:"running"
  });

}



//=========================
// POST REQUESTS
//=========================

function doPost(e){

  const data = JSON.parse(e.postData.contents);


  //--------------------------------
  // LOGIN
  //--------------------------------

  if(data.action=="login"){

    if(data.code==AUTH_CODE){

      return json({
        status:"success"
      });

    }

    return json({
      status:"invalid",
      message:"Wrong authentication code."
    });

  }



  //--------------------------------
  // MEAL SCAN
  //--------------------------------

  const guestId=data.guestId;
  const mealSlot=data.mealSlot;
  const station=data.station||"";

  const sheet=SpreadsheetApp.getActive().getSheetByName(SHEET_GUESTS);

  const values=sheet.getDataRange().getValues();

  const headers=values[0];

  const mealColumn=headers.indexOf(mealSlot);

  if(mealColumn==-1){

    return json({
      status:"error",
      message:"Meal slot not found."
    });

  }


  for(let i=1;i<values.length;i++){

    if(values[i][0].toString().trim().toUpperCase()==guestId){

      // Already Taken

      if(values[i][mealColumn]==true){

        return json({

          status:"duplicate",

          name:values[i][1],

          role:values[i][2],

          mealLabel:mealSlot,

          mealsCompleted:countMeals(values[i],headers),

          totalSlots:headers.length-3,

          message:"Meal already collected."

        });

      }


      // Mark Meal

      sheet.getRange(i+1,mealColumn+1).setValue(true);


      SpreadsheetApp.getActive()
      .getSheetByName(SHEET_LOGS)
      .appendRow([

        new Date(),

        guestId,

        values[i][1],

        mealSlot,

        station

      ]);


      return json({

        status:"success",

        name:values[i][1],

        role:values[i][2],

        mealLabel:mealSlot,

        mealsCompleted:countMealsUpdated(values[i],headers),

        totalSlots:headers.length-3

      });

    }

  }


  return json({

    status:"invalid",

    message:"Guest not found."

  });

}



//=========================
// HELPERS
//=========================

function countMeals(row,headers){

  let c=0;

  for(let i=3;i<headers.length;i++){

    if(row[i]==true)
      c++;

  }

  return c;

}


function countMealsUpdated(row,headers){

  let c=1;

  for(let i=3;i<headers.length;i++){

    if(row[i]==true)
      c++;

  }

  return c;

}



function json(obj){

  return ContentService
  .createTextOutput(JSON.stringify(obj))
  .setMimeType(ContentService.MimeType.JSON);

}