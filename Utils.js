/*======================================================
  UTILITY MODULE
======================================================*/


function utilsRandom(length){
  return utilsRandomString(length);

}


/*======================================================
  GENERATE LICENSE
======================================================*/
function utilsGenerateLicense(){

  return "SSB-"

      +utilsRandom(4)

      +"-"

      +utilsRandom(4)

      +"-"

      +utilsRandom(4);

}


/*======================================================
  CUSTOMER ID
======================================================*/
function utilsGenerateCustomerID(){

  return "CUS-"

      +Utilities.formatDate(

        new Date(),

        Session.getScriptTimeZone(),

        "yyMMddHHmmss"

      );

}


/*======================================================
  ENTERPRISE LICENSE UTILITIES
======================================================*/

function utilsRandomString(length){

  const chars="ABCDEFGHJKLMNPQRSTUVWXYZ23456789";

  let value="";

  while(value.length<length){

    value+=Utilities.getUuid().replace(/-/g,"").toUpperCase();

  }

  return value.substring(0,length).split("").map(function(char){

    return chars.charAt(char.charCodeAt(0)%chars.length);

  }).join("");

}


function utilsGenerateSecureLicenseKey(){

  return "SSB-"+utilsRandomString(4)+"-"+utilsRandomString(4)+"-"+utilsRandomString(4);

}


function utilsNextLicenseId(sheet){

  const year=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy");

  return utilsNextSequentialId_(sheet,"LIC-"+year+"-",6);

}


function utilsNextCustomerId(sheet){

  return utilsNextSequentialId_(sheet,"CUS-",6);

}


function utilsNextSequentialId_(sheet,prefix,width){

  const values=sheet.getLastRow()>1
    ?sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues()
    :[];

  let highest=0;

  values.forEach(function(row){

    const value=String(row[0]);

    if(value.indexOf(prefix)!==0) return;

    const number=Number(value.substring(prefix.length));

    if(Number.isInteger(number)) highest=Math.max(highest,number);

  });

  return prefix+String(highest+1).padStart(width,"0");

}


function utilsGetRequiredHeaderMap(sheet,requiredFields){

  if(sheet.getLastRow()<1){

    throw new Error('Sheet "'+sheet.getName()+'" is empty.');

  }

  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0]
    .map(function(header){return String(header).trim().toLowerCase();});

  const map={};

  requiredFields.forEach(function(field){

    const index=headers.findIndex(function(header){

      return field.aliases.indexOf(header)!==-1;

    });

    if(index===-1){

      throw new Error('Sheet "'+sheet.getName()+'" is missing required header "'+field.label+'".');

    }

    map[field.key]=index+1;

  });

  return map;

}
