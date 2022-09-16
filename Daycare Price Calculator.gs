var formSheet = SpreadsheetApp.openById("/*Insert ID Here*/");
var email, hasChildInDaycare4, employeeName, childInDaycare4Name, child2Name, child2MinutesPerWeek, child3Name, child3MinutesPerWeek;
const _PRICE_PER_HOUR_FIRST_CHILD = 6.50;
const _PRICE_PER_HOUR_AFTER_FIRST_CHILD = 5.50;
const _FLAT_RATE_FOR_DAYCARE_4 = 3800.00;
const _WEEKS_PER_YEAR = 34;
const _CHILD_2_FIRST_HOUR_COL = 13;
const _CHILD_2_LAST_HOUR_COL = 27;
const _CHILD_3_FIRST_HOUR_COL = 33;
const _CHILD_3_LAST_HOUR_COL = 47;
const dayOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
const currency = Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
});
var numErrors = 0;
var errorMsg = "<ol>";
var numOfChildren = 0;
var child1Price = 0;
var child2Price = 0;
var child3Price = 0;
var emailRow = 0;
var latestEntry = [];

function manualSend() {
  var rowToSend = 62;
  emailRow = rowToSend;
  var sheetData = formSheet.getSheets()[0].getDataRange().getValues();
  for (var col = 0; col < formSheet.getLastColumn(); col++) {
      latestEntry.push(sheetData[rowToSend - 1][col]);
    }
  var e = {values: latestEntry};
  // main(e);
}

function main(e) {
  // ScriptApp.newTrigger("main").forSpreadsheet(formSheet).onFormSubmit().create();
  
  latestEntry = e.values;
  for (var i in latestEntry) {
    if ((i >= _CHILD_2_FIRST_HOUR_COL && i <= _CHILD_2_LAST_HOUR_COL) || (i >= _CHILD_3_FIRST_HOUR_COL && i <= _CHILD_3_LAST_HOUR_COL)){
      if (latestEntry[i] != "" && latestEntry[i] != "Yes" && latestEntry[i] != "No") {
        if (typeof latestEntry[i] == "object") {
          latestEntry[i] = new Time(getTimeStr(latestEntry[i]));
        }
        else {
          latestEntry[i] = new Time(latestEntry[i]);
        }       
        console.log(`${i}: ${latestEntry[i]}`);
        continue;
      }
      else {
        console.log(`${i}: ${latestEntry[i]}`);
        continue;
      }
    }
    console.log(`${i}: ${latestEntry[i]}`);
  }
  email = latestEntry[1];

employeeName = `${latestEntry[3]} ${latestEntry[2]}`;
if (latestEntry[5] == "Yes") {
  hasChildInDaycare4 = true;
  childInDaycare4Name = `${latestEntry[7]} ${latestEntry[6]}`;
  numOfChildren++;
}
else {
  hasChildInDaycare4 = false;
}
if (latestEntry[9] == "Yes") {
child2Name = `${latestEntry[11]} ${latestEntry[10]}`;
checkforErrors(_CHILD_2_FIRST_HOUR_COL, _CHILD_2_LAST_HOUR_COL);
if (numErrors == 0) {
child2MinutesPerWeek = calcTotalMinutes(_CHILD_2_FIRST_HOUR_COL, _CHILD_2_LAST_HOUR_COL);}
numOfChildren++;
}
if (latestEntry[28] == "Yes") {
child3Name = `${latestEntry[30]} ${latestEntry[29]}`;
checkforErrors(_CHILD_3_FIRST_HOUR_COL, _CHILD_3_LAST_HOUR_COL);
if (numErrors == 0) {
latestEntry[32] == "Same drop off and pick up times" ? child3MinutesPerWeek = child2MinutesPerWeek : child3MinutesPerWeek = calcTotalMinutes(_CHILD_3_FIRST_HOUR_COL, _CHILD_3_LAST_HOUR_COL);}
numOfChildren++;
}
if (emailRow == 0) {emailRow = formSheet.getSheets()[0].getLastRow();}
if (numErrors != 0) {
  emailError();
  return;
}
calcPrice();
emailPrice();
}

function checkforErrors(beg, end) {
  var day = 0;
  for (var i = beg; i < end; i += 3) {
    day++;
    beg < 30 ? childName = child2Name : childName = child3Name;
    if (latestEntry[i] == "No") {
      if (latestEntry[i + 1] != "") {
        errorMsg += `<li>You selected that ${childName} will not be in daycare on ${dayOfWeek[day]}, but you entered that they will be dropped off at ${latestEntry[i + 1]}</li><br />`;
        numErrors++;
      }
      if (latestEntry[i + 2] != "") {
      errorMsg += `<li>You selected that ${childName} will not be in daycare on ${dayOfWeek[day]}, but you entered that they will be picked up at ${latestEntry[i + 2]}</li><br />`;
      numErrors++;
      }
    }    
     if (latestEntry[i] == "Yes") {
       if (latestEntry[i + 1] == "") {
         errorMsg += `<li>You selected that ${childName} will be in daycare on ${dayOfWeek[day]}, but you did not enter a drop off time</li><br />`;
         numErrors++;
       }
       if (latestEntry[i + 2] == "") {
         errorMsg += `<li>You selected that ${childName} will be in daycare on ${dayOfWeek[day]}, but you did not enter a pick up time</li><br />`;
         numErrors++;
       }
     }
     try {
       if ((latestEntry[i + 1].getTime() > latestEntry[i + 2].getTime())) {     
       errorMsg += `<li>The pick up time for ${childName} on ${dayOfWeek[day]} (${latestEntry[i + 2]}) is not later than the drop off time (${latestEntry[i + 1]})</li><br />`;
       numErrors++;
     }
     } catch (TypeError){}
     try {
    if (latestEntry[i + 1].getTime() < 510) {
      errorMsg += `<li>The drop off time for ${childName} on ${dayOfWeek[day]} (${latestEntry[i + 1]}) is earlier than 8:30 AM</li><br />`;
      numErrors++;
    }
     } catch (TypeError){}
    try {
    if (day != 5 && latestEntry[i + 2].getTime() > 1050) {
      errorMsg += `<li>The pick up time for ${childName} on ${dayOfWeek[day]} (${latestEntry[i + 2]}) is later than 5:30 PM</li><br />`;
      numErrors++;
    }
    } catch (TypeError){}
    try {
    if (day == 5 && latestEntry[i + 2].getTime() >= 780) {
      errorMsg += `<li>The pick up time for ${childName} on ${dayOfWeek[day]} (${latestEntry[i + 2]}) is later than 1:00 PM</li><br />`;
      numErrors++;
    }
    } catch (TypeError){}
  }
  errorMsg += "</ol>";
}

function calcTotalMinutes(beg, end) {
  var totalMinutesPerWeek = 0;
  for (var i = beg; i < end; i += 3) {
  if (latestEntry[i] == "No") {continue;}  
    totalMinutesPerWeek += (latestEntry[i + 2].getTime() - latestEntry[i + 1].getTime());
  }  
  return totalMinutesPerWeek;
}

function calcPrice() {
  if (numOfChildren == 1 && hasChildInDaycare4) {
    child1Price = _FLAT_RATE_FOR_DAYCARE_4;
  }
  if (numOfChildren == 1 && !hasChildInDaycare4) {
    child2Price = child2MinutesPerWeek * (_PRICE_PER_HOUR_FIRST_CHILD / 60);
  }
  if (numOfChildren == 2 && hasChildInDaycare4) {
    child1Price = _FLAT_RATE_FOR_DAYCARE_4;
    child2Price = child2MinutesPerWeek * (_PRICE_PER_HOUR_AFTER_FIRST_CHILD / 60);
  }
  if (numOfChildren == 2 && !hasChildInDaycare4) {
    child2Price = child2MinutesPerWeek * (_PRICE_PER_HOUR_FIRST_CHILD / 60);
    child3Price = child3MinutesPerWeek * (_PRICE_PER_HOUR_AFTER_FIRST_CHILD / 60);    
  }
  if (numOfChildren == 3 && hasChildInDaycare4) {
    child1Price = _FLAT_RATE_FOR_DAYCARE_4;
    child2Price = child2MinutesPerWeek * (_PRICE_PER_HOUR_AFTER_FIRST_CHILD / 60);
    child3Price = child3MinutesPerWeek * (_PRICE_PER_HOUR_AFTER_FIRST_CHILD / 60);    
  }
}

function emailPrice() {
  var child2Time, child3Time ;
  child2MinutesPerWeek % 60 != 0 ? child2Time = `${Math.floor(child2MinutesPerWeek / 60)} hrs ${child2MinutesPerWeek % 60} min` : child2Time = `${Math.floor(child2MinutesPerWeek / 60)} hrs`;
  child3MinutesPerWeek % 60 != 0 ? child3Time = `${Math.floor(child3MinutesPerWeek / 60)} hrs ${child3MinutesPerWeek % 60} min` : child3Time = `${Math.floor(child3MinutesPerWeek / 60)} hrs`;
  var body = `Bnos Yisroel Daycare for ${employeeName}:` + "\n";
  if (child1Price != 0) {
    body += `${childInDaycare4Name}: ${currency.format(child1Price)} (Daycare #4 Fixed rate)` + "\n";
  }
  if (child2Price != 0 && !hasChildInDaycare4) {
    body += `${child2Name}: ${child2Time} x ${currency.format(_PRICE_PER_HOUR_FIRST_CHILD)}/hr for ${_WEEKS_PER_YEAR} weeks = ${currency.format(child2Price * _WEEKS_PER_YEAR)}` + "\n";
  }
  if (child2Price != 0 && hasChildInDaycare4) {
    body += `${child2Name}: ${child2Time} x ${currency.format(_PRICE_PER_HOUR_AFTER_FIRST_CHILD)}/hr for ${_WEEKS_PER_YEAR} weeks = ${currency.format(child2Price * _WEEKS_PER_YEAR)}` + "\n";
  }
  if (child3Price != 0 ) {
    body += `${child3Name}: ${child3Time} x ${currency.format(_PRICE_PER_HOUR_AFTER_FIRST_CHILD)}/hr for ${_WEEKS_PER_YEAR} weeks = ${currency.format(child3Price * _WEEKS_PER_YEAR)}` + "\n";
  }
  body += "------------------------------------------------\n";
  body += `Total: ${currency.format(child1Price + (child2Price * _WEEKS_PER_YEAR) + (child3Price * _WEEKS_PER_YEAR))}`;
  console.log(body);
  GmailApp.sendEmail(email, "Bnos Yisroel Daycare", body, {
    name: "Bnos Yisroel",
    bcc: "EMAIL REMOVED"
  });
  console.log(`Remaining emails: ${MailApp.getRemainingDailyQuota()}`);
  formSheet.getRange("AW" + emailRow).setValue(body);
}

function getTimeStr(dateObj) {
  var hours = dateObj.getHours();
  var amOrPm = hours >= 12 ? 'PM' : 'AM';
  hours % 12 == 0 ? hours = 12 : hours = Time.addLeadingZeroIfNone(hours % 12);
  return `${hours}:${Time.addLeadingZeroIfNone(dateObj.getMinutes())}:00 ${amOrPm}`;
}

 String.prototype.getFormTime = function () {
    return "";
};

 String.prototype.getFormDate = function () {
   var str = this.valueOf();
    if (str == "") {return "";}
    else {
      str = str.split("/");
      return `${str[2]}-${Time.addLeadingZeroIfNone(str[1])}-${Time.addLeadingZeroIfNone(str[0])}`;
    }
};

 Date.prototype.getFormDate = function () {
   return `${this.getFullYear()}-${Time.addLeadingZeroIfNone(this.getMonth() + 1)}-${Time.addLeadingZeroIfNone(this.getDate())}`;
 }

  function emailError() {
    var begStr;
    numErrors == 1 ? begStr = "was 1 error" : begStr = `were ${numErrors} errors`;
    var html = `<p>There ${begStr} with your form: <br />${errorMsg}Please fill out the form again: <a href="https://docs.google.com/forms/d/e/1FAIpQLSfsKSuEru2eo2s06ODZxxr2VVzOM13V1fPBjk0HMXHgIFWTmQ/viewform?usp=pp_url&entry.1800913134=${latestEntry[1]}&entry.1042141425=${latestEntry[2]}&entry.480052026=${latestEntry[3]}&entry.33757732=${latestEntry[4]}&entry.1244777689=${latestEntry[5]}&entry.1000839449=${latestEntry[6]}&entry.82263155=${latestEntry[7]}&entry.968581035=${latestEntry[8].getFormDate()}&entry.1235376051=${latestEntry[9]}&entry.1378877323=${latestEntry[10]}&entry.1711419373=${latestEntry[11]}&entry.945587762=${latestEntry[12].getFormDate()}&entry.1388499204=${latestEntry[13]}&entry.1751498840=${latestEntry[14].getFormTime()}&entry.638022772=${latestEntry[15].getFormTime()}&entry.277205963=${latestEntry[16]}&entry.1823075640=${latestEntry[17].getFormTime()}&entry.827963157=${latestEntry[18].getFormTime()}&entry.1374580991=${latestEntry[19]}&entry.1517967336=${latestEntry[20].getFormTime()}&entry.773110373=${latestEntry[21].getFormTime()}&entry.951892293=${latestEntry[22]}&entry.84413758=${latestEntry[23].getFormTime()}&entry.1848730221=${latestEntry[24].getFormTime()}&entry.161735636=${latestEntry[25]}&entry.886445615=${latestEntry[26].getFormTime()}&entry.538688772=${latestEntry[27].getFormTime()}&entry.286811188=${latestEntry[28]}&entry.1202009627=${latestEntry[29]}&entry.1603674758=${latestEntry[30]}&entry.645693449=${latestEntry[31].getFormDate()}&entry.678308094=${latestEntry[32]}&entry.1800350375=${latestEntry[33]}&entry.1041551511=${latestEntry[34].getFormTime()}&entry.1287864189=${latestEntry[35].getFormTime()}&entry.1628039415=${latestEntry[36]}&entry.1131456671=${latestEntry[37].getFormTime()}&entry.146137624=${latestEntry[38].getFormTime()}&entry.1185188487=${latestEntry[39]}&entry.1390395854=${latestEntry[40].getFormTime()}&entry.2027872313=${latestEntry[41].getFormTime()}&entry.1817004232=${latestEntry[42]}&entry.1363729082=${latestEntry[43].getFormTime()}&entry.824744781=${latestEntry[44].getFormTime()}&entry.954951242=${latestEntry[45]}&entry.470222142=${latestEntry[46].getFormTime()}&entry.265443800=${latestEntry[47].getFormTime()}">https://docs.google.com/forms/d/e/1FAIpQLSfsKSuEru2eo2s06ODZxxr2VVzOM13V1fPBjk0HMXHgIFWTmQ/viewform</a><br />Thank you!</p>`;
    GmailApp.sendEmail(email, "Bnos Yisroel Daycare", "", {
      name: "Bnos Yisroel",
      htmlBody: html,
      bcc: "EMAIL REMOVED"
    });
    console.log("Sent: " + html);
    console.log(`Remaining emails: ${MailApp.getRemainingDailyQuota()}`);
    formSheet.getRange("AW" + emailRow).setValue(html);
  }
