/*======================================================
  UTILITY MODULE
======================================================*/


function utilsRandom(length){

  const chars=
  "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";

  let text="";

  for(let i=0;i<length;i++){

    text+=chars.charAt(
      Math.floor(
        Math.random()*chars.length
      )
    );

  }

  return text;

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