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

  const chart={

    license:dashboardBuildMonthlySeries_(shLicense,totalLicense),

    customer:dashboardBuildMonthlySeries_(shCustomer,totalCustomer),

    status:{

      active:totalActive,

      expired:totalExpired

    }

  };

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

    recentActivity:recentActivity,

    chart:chart

  };

}

/*======================================================
  DASHBOARD DATA
======================================================*/

function dashboardBuildMonthlySeries_(sheet,currentTotal){

  const timeZone=Session.getScriptTimeZone();

  const currentMonth=Utilities.formatDate(

    new Date(),

    timeZone,

    "MMM yyyy"

  );

  const fallback={

    labels:[currentMonth],

    values:[currentTotal]

  };

  if(sheet.getLastRow()<2) return fallback;

  const rows=sheet.getDataRange().getValues();

  const headers=rows[0].map(function(value){

    return String(value).trim().toLowerCase();

  });

  const dateIndex=headers.findIndex(function(header){

    return /date|tanggal|created|registered|issued/.test(header);

  });

  if(dateIndex===-1) return fallback;

  const months={};

  rows.slice(1).forEach(function(row){

    const value=row[dateIndex];

    const date=value instanceof Date?value:new Date(value);

    if(isNaN(date.getTime())) return;

    const key=Utilities.formatDate(date,timeZone,"yyyy-MM");

    if(!months[key]){

      months[key]={

        label:Utilities.formatDate(date,timeZone,"MMM yyyy"),

        value:0

      };

    }

    months[key].value++;

  });

  const keys=Object.keys(months).sort().slice(-12);

  if(keys.length===0) return fallback;

  return{

    labels:keys.map(function(key){

      return months[key].label;

    }),

    values:keys.map(function(key){

      return months[key].value;

    })

  };

}
