/*=========================================================
  SSB Attendance Pro Enterprise
  Module : Dashboard
  Version : 2.1.0
=========================================================*/


/*=========================================================
  DASHBOARD DATA
=========================================================*/
function dashboardGetData(){

  const ss = SpreadsheetApp.getActive();

  const shLicense = ss.getSheetByName("LICENSE");
  const shCustomer = ss.getSheetByName("CUSTOMER");
  const shLog = ss.getSheetByName("LOG");
  const shSetting = ss.getSheetByName("SETTING");

  // =========================
  // LICENSE
  // =========================

  const totalLicense = Math.max(shLicense.getLastRow()-1,0);

  // =========================
  // CUSTOMER
  // =========================

  const totalCustomer = Math.max(shCustomer.getLastRow()-1,0);

  // =========================
  // ACTIVE & EXPIRED
  // =========================

  let totalActive = 0;
  let totalExpired = 0;

  if(totalLicense>0){

    const status = shLicense
      .getRange(2,3,totalLicense,1)
      .getValues();

    status.forEach(function(r){

      if(r[0]=="Active") totalActive++;

      if(r[0]=="Expired") totalExpired++;

    });

  }

  // =========================
  // RECENT ACTIVITY
  // =========================

  let recentActivity=[];

  const lastLog=shLog.getLastRow();

  if(lastLog>1){

    const logData=shLog
      .getRange(
        Math.max(2,lastLog-4),
        1,
        Math.min(5,lastLog-1),
        3
      )
      .getValues();

    logData.reverse();

    logData.forEach(function(r){

      recentActivity.push({

        date:r[0],

        action:r[1],

        detail:r[2]

      });

    });

  }

  // =========================
  // VERSION
  // =========================

  let version="2.1.0";

  if(shSetting.getLastRow()>1){

    const data=shSetting.getRange(2,1,shSetting.getLastRow()-1,2).getValues();

    data.forEach(function(r){

      if(r[0]=="Version"){

        version=r[1];

      }

    });

  }

  // =========================
  // RETURN
  // =========================

  return{

    totalLicense:totalLicense,

    totalCustomer:totalCustomer,

    totalActive:totalActive,

    totalExpired:totalExpired,

    version:version,

    database:"Connected",

    server:"Offline",

    recentActivity:recentActivity

  };

}
