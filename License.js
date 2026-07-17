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

function activationGetState(){
  try{
    const ss=SpreadsheetApp.getActive();
    const sheet=activationEnsureSystemSheet_(ss);
    const values=sheet.getDataRange().getValues();
    const state={};

    values.slice(1).forEach(function(row){
      if(!row || !String(row[0]||"").trim()) return;
      state[String(row[0]).trim().toLowerCase()]=String(row[1]||"").trim();
    });

    const activated=state.activated==="true";
    const licenseKey=String(state.licensekey||"").trim();
    const spreadsheetId=String(state.spreadsheetid||"").trim();
    const activatedOn=state.activatedon||"-";
    const expiredOn=state.expiry||"-";

    if(!licenseKey){
      return{activated:false,message:"Your software has not been activated yet. Please enter a valid Enterprise License Key.",licenseStatus:"Pending",bindingStatus:"Unbound",spreadsheetId:"-",activatedOn:"-",expiredOn:"-",database:"Google Sheets"};
    }

    const validation=activationValidateLicenseRecord_(ss,licenseKey);
    if(!validation.valid){
      activationClearSystemState_(ss);
      return{activated:false,message:validation.message||"Activation state is invalid.",licenseStatus:"Invalid",bindingStatus:"Unbound",spreadsheetId:"-",activatedOn:"-",expiredOn:"-",database:"Google Sheets"};
    }

    if(spreadsheetId && spreadsheetId!==ss.getId()){
      activationClearSystemState_(ss);
      return{activated:false,message:"This license is bound to another spreadsheet. Please activate it on the original installation.",licenseStatus:"Invalid",bindingStatus:"Already Bound",spreadsheetId:spreadsheetId,activatedOn:activatedOn,expiredOn:expiredOn,database:"Google Sheets"};
    }

    activationPersistSystemState_(ss,licenseKey,ss.getId(),validation.version,validation.expiryDisplay,validation.status);
    return{activated:true,message:"Enterprise activation is active.",licenseKey:licenseKey,licenseStatus:validation.status||"Active",bindingStatus:"Bound",spreadsheetId:ss.getId(),activatedOn:activatedOn||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd"),expiredOn:validation.expiryDisplay||"-",database:"Google Sheets",version:validation.version};

  }catch(err){
    return{activated:false,message:err&&err.message?err.message:"Unable to validate activation state."};
  }
}

function enterpriseActivate(payload){
  try{
    const licenseKey=String(payload&&payload.licenseKey||"").trim();
    if(!licenseKey){
      return{success:false,message:"License key is required.",statusType:"warning"};
    }

    const ss=SpreadsheetApp.getActive();
    const licenseSheet=ss.getSheetByName("LICENSE");
    if(!licenseSheet){
      return{success:false,message:"LICENSE sheet is required.",statusType:"danger",licenseStatus:"Invalid",bindingStatus:"Unbound",spreadsheetId:"-",activatedOn:"-",expiredOn:"-",database:"Google Sheets"};
    }

    const validation=activationValidateLicenseRecord_(ss,licenseKey);
    if(!validation.valid){
      activationClearSystemState_(ss);
      return{success:false,message:validation.message||"Invalid license key.",statusType:"danger",licenseStatus:"Invalid",bindingStatus:"Unbound",spreadsheetId:"-",activatedOn:"-",expiredOn:"-",database:"Google Sheets"};
    }

    if(validation.boundSpreadsheetId && validation.boundSpreadsheetId!==ss.getId()){
      activationClearSystemState_(ss);
      return{success:false,message:"License already used on another spreadsheet.",statusType:"danger",licenseStatus:"Invalid",bindingStatus:"Already Bound",spreadsheetId:validation.boundSpreadsheetId,activatedOn:"-",expiredOn:validation.expiryDisplay||"-",database:"Google Sheets"};
    }

    if(!validation.boundSpreadsheetId){
      licenseSheet.getRange(validation.rowIndex+1,8).setValue(ss.getId());
    }

    activationPersistSystemState_(ss,licenseKey,ss.getId(),validation.version,validation.expiryDisplay,validation.status);
    return{success:true,activated:true,message:"Enterprise activation completed successfully.",statusType:"success",licenseStatus:validation.status||"Active",bindingStatus:"Bound",spreadsheetId:ss.getId(),activatedOn:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd"),expiredOn:validation.expiryDisplay||"-",database:"Google Sheets",version:validation.version};

  }catch(err){
    return{success:false,message:err&&err.message?err.message:"Activation could not be completed.",statusType:"danger"};
  }
}

function activationValidateLicenseRecord_(ss,licenseKey){
  const licenseSheet=ss.getSheetByName("LICENSE");
  if(!licenseSheet){
    throw new Error("LICENSE sheet is required.");
  }

  const values=licenseSheet.getDataRange().getValues();
  let match=null;

  values.forEach(function(row,index){
    if(!row || !String(row[0]||"").trim()) return;
    if(String(row[0]).trim()!==licenseKey) return;
    match={rowIndex:index,row:row};
  });

  if(!match){
    return{valid:false,message:"License key is invalid."};
  }

  const row=match.row;
  const status=String(row[2]||"").trim();
  const version=String(row[6]||"").trim();
  const expiryValue=row[3];
  const boundSpreadsheetId=String(row[7]||"").trim();
  const currentDate=new Date();
  let expiryDate=null;

  if(expiryValue instanceof Date){
    expiryDate=expiryValue;
  }else if(typeof expiryValue==="string"){
    expiryDate=new Date(expiryValue);
  }else if(expiryValue){
    expiryDate=new Date(expiryValue);
  }

  if(!status || status.toLowerCase()!=="active"){
    return{valid:false,message:"License is not active."};
  }

  if(!version){
    return{valid:false,message:"License version is invalid."};
  }

  if(expiryDate && expiryDate.getTime()<currentDate.getTime()){
    return{valid:false,message:"License expired."};
  }

  return{
    valid:true,
    rowIndex:match.rowIndex,
    boundSpreadsheetId:boundSpreadsheetId,
    version:version,
    status:status,
    expiryDisplay:expiryDate?Utilities.formatDate(expiryDate,Session.getScriptTimeZone(),"yyyy-MM-dd"):(row[3]?String(row[3]).trim():"")
  };
}

function activationEnsureSystemSheet_(ss){
  let sheet=ss.getSheetByName("SYSTEM");

  if(!sheet){
    sheet=ss.insertSheet("SYSTEM");
  }

  const values=sheet.getDataRange().getValues();
  if(values.length===0 || values[0].length===0){
    sheet.getRange(1,1,1,2).setValues([["Key","Value"]]);
  }

  return sheet;
}

function activationPersistSystemState_(ss,licenseKey,spreadsheetId,version,expiry,status){
  const sheet=activationEnsureSystemSheet_(ss);
  const activatedOn=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd");
  activationSetSystemValue_(sheet,"activated","true");
  activationSetSystemValue_(sheet,"licenseKey",licenseKey);
  activationSetSystemValue_(sheet,"spreadsheetId",spreadsheetId);
  activationSetSystemValue_(sheet,"version",version);
  activationSetSystemValue_(sheet,"expiry",expiry||"");
  activationSetSystemValue_(sheet,"status",status||"");
  activationSetSystemValue_(sheet,"activatedOn",activatedOn);
}

function activationClearSystemState_(ss){
  const sheet=activationEnsureSystemSheet_(ss);
  activationSetSystemValue_(sheet,"activated","false");
  activationSetSystemValue_(sheet,"licenseKey","");
  activationSetSystemValue_(sheet,"spreadsheetId","");
  activationSetSystemValue_(sheet,"version","");
  activationSetSystemValue_(sheet,"expiry","");
  activationSetSystemValue_(sheet,"status","");
  activationSetSystemValue_(sheet,"activatedOn","");
}

function activationSetSystemValue_(sheet,key,value){
  const values=sheet.getDataRange().getValues();
  const trimmedKey=String(key||"").trim();
  if(!trimmedKey){
    return;
  }

  let foundIndex=-1;
  values.forEach(function(row,index){
    if(index===0) return;
    if(String(row[0]||"").trim().toLowerCase()===trimmedKey.toLowerCase()){
      foundIndex=index;
    }
  });

  if(foundIndex===-1){
    sheet.appendRow([trimmedKey,value]);
    return;
  }

  sheet.getRange(foundIndex+1,2).setValue(value);
}
