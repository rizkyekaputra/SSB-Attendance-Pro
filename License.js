function licenseGenerate(payload){
  const data=licenseValidatePayload_(payload);
  const ss=SpreadsheetApp.getActive();
  const licenses=ss.getSheetByName("LICENSE");
  const customers=ss.getSheetByName("CUSTOMER");
  const logs=ss.getSheetByName("LOG");
  if(!licenses||!customers||!logs) throw new Error("LICENSE, CUSTOMER, and LOG sheets are required.");
  const licenseKey=licenseGenerateUniqueKey_(licenses);
  const customer=customerFindOrCreate_(customers,data,licenseKey);
  const licenseId=utilsNextLicenseId(licenses);
  const createdAt=new Date(),expiredAt=new Date(createdAt);
  expiredAt.setMonth(expiredAt.getMonth()+data.duration);
  licenses.appendRow([licenseKey,data.ssbName,"Active",expiredAt,data.email,createdAt,data.version,ss.getId(),data.note]);
  const expiry=Utilities.formatDate(expiredAt,Session.getScriptTimeZone(),"yyyy-MM-dd");
  logs.appendRow([createdAt,"Generate License","License ID: "+licenseId+" | License Key: "+licenseKey+" | Nama SSB: "+data.ssbName+" | Customer: "+data.customerName+" | Version: "+data.version+" | Expired: "+expiry+" | Administrator: Admin"]);
  return{success:true,licenseId:licenseId,licenseKey:licenseKey,customerId:customer.id,expiredAt:expiry};
}
function customerFindOrCreate_(sheet,data,licenseKey){
  const rows=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,8).getValues():[];
  const found=rows.find(function(row){return String(row[1]).trim()===data.ssbName&&String(row[3]).trim()===data.whatsapp&&String(row[4]).trim().toLowerCase()===data.email.toLowerCase();});
  if(found) return{id:found[0]};
  const id=utilsNextCustomerId(sheet);
  sheet.appendRow([id,data.ssbName,data.customerName,data.whatsapp,data.email,data.city,new Date(),licenseKey]);
  return{id:id};
}
function licenseGenerateUniqueKey_(sheet){
  const keys=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().flat():[];
  for(let i=0;i<10;i++){const key=utilsGenerateSecureLicenseKey();if(keys.indexOf(key)===-1)return key;}
  throw new Error("Unable to generate a unique license key.");
}
function licenseValidatePayload_(payload){
  const fields=["ssbName","customerName","whatsapp","email","city","version","duration"];
  fields.forEach(function(key){if(!payload||!String(payload[key]||"").trim())throw new Error("Field "+key+" is required.");});
  if(!/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/.test(payload.email))throw new Error("Email format is invalid.");
  if(!/^[0-9+()\\-\\s]{8,20}$/.test(payload.whatsapp))throw new Error("Phone number format is invalid.");
  const duration=Number(payload.duration);if(!Number.isInteger(duration)||duration<1)throw new Error("Expired Duration must be a positive whole number of months.");
  return{ssbName:String(payload.ssbName).trim(),customerName:String(payload.customerName).trim(),whatsapp:String(payload.whatsapp).trim(),email:String(payload.email).trim(),city:String(payload.city).trim(),version:String(payload.version).trim(),duration:duration,note:String(payload.note||"").trim()};
}
