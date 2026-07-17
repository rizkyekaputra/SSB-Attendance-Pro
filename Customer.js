function customerFormatDateForDisplay_(value){

  if(!value) return "";

  const rawValue=String(value).trim();
  if(!rawValue) return "";

  const parsedDate=new Date(rawValue);

  if(isNaN(parsedDate.getTime())) return rawValue;

  return Utilities.formatDate(parsedDate,Session.getScriptTimeZone(),"dd/MM/yyyy");

}

function customerGetData(){

  const ss=SpreadsheetApp.getActive();
  const sheet=ss.getSheetByName("CUSTOMER");

  if(!sheet){

    throw new Error("CUSTOMER sheet is required.");

  }

  const values=sheet.getDataRange().getDisplayValues();
  const customers=[];
  let withLicense=0;
  let thisMonth=0;

  const today=new Date();
  const currentMonth=Utilities.formatDate(today,Session.getScriptTimeZone(),"yyyy-MM");

  for(let i=1;i<values.length;i++){

    const row=values[i];

    if(!row || row.every(function(value){return String(value).trim()==="";})) continue;

    const city=String(row[5]||"").trim();
    const purchaseDate=String(row[6]||"").trim();
    const license=String(row[7]||"").trim();

    if(license) withLicense++;

    if(purchaseDate){

      const parsedDate=new Date(purchaseDate);
      const monthKey=Utilities.formatDate(parsedDate,Session.getScriptTimeZone(),"yyyy-MM");

      if(!isNaN(parsedDate.getTime()) && monthKey===currentMonth) thisMonth++;

    }

    customers.push({

      id:String(row[0]||"").trim(),
      ssbName:String(row[1]||"").trim(),
      buyerName:String(row[2]||"").trim(),
      whatsapp:String(row[3]||"").trim(),
      email:String(row[4]||"").trim(),
      city:city,
      purchaseDate:customerFormatDateForDisplay_(purchaseDate),
      license:license

    });

  }

  return{

    customers:customers,
    stats:{
      total:customers.length,
      withLicense:withLicense,
      thisMonth:thisMonth,
      hasData:customers.length>0
    }

  };

}
