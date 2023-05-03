/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, setInterval, clearInterval, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady().then(function () {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  Excel.run(async (context) => {
    eventsOn();
    let scheduleSheet = context.workbook.worksheets.getItem("Schedule");
    let driverSheet = context.workbook.worksheets.getItem("Drivers");
    let startTimeSheet = context.workbook.worksheets.getItem("Schedule By Start Time");
    let tucSheet = context.workbook.worksheets.getItem("TUC");
    let ducSheet = context.workbook.worksheets.getItem("DUC");
    updateStatus();
    let cycles = 0;
    let startRange = startTimeSheet.getRange("A1");
    let cycleRange = startTimeSheet.getRange("F1");
    let intervalID = setInterval(digitalClock, 300000);
    cycleRange.values = cycles as any;
    startRange.values = intervalID as any;
    await context.sync();
    scheduleSheet.onChanged.add(onScheduleChange);
    driverSheet.onChanged.add(onDriverChange);
    startTimeSheet.onChanged.add(onStartTimeChange);
    tucSheet.onSelectionChanged.add(onTUCSelection);
    ducSheet.onSelectionChanged.add(onDUCSelection);
    await context.sync();
  });
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();
      // Read the range address
      range.load("address");
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

async function onScheduleChange(event: Excel.WorksheetChangedEventArgs) {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    if (event.address.substring(0, 1) !== "I" || event.triggerSource === "ThisLocalAddin") {
      return;
    }
    let sheet = context.workbook.worksheets.getItem("Schedule");
    let range = sheet.getRange("I103:I200");
    //sheet.protection.pauseProtection("CPM");
    const sortFields = [{ key: 0, ascending: true }];
    range.sort.apply(sortFields);
    //sheet.protection.resumeProtection();
    await context.sync();
  });
}

async function onDriverChange(event: Excel.WorksheetChangedEventArgs) {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    try {
      if (event.triggerSource === "ThisLocalAddin") {
        return;
      }
      if (event.address.includes(":")) {
        deleteRow();
        return;
      }
      let sheet = context.workbook.worksheets.getItem("Drivers");
      let range = sheet.getRange(event.address);
      let details = event.details;
      let values = details.valueAfter;
      let col = event.address.substring(0, 1);
      switch (col) {
        case "A":
        case "D":
        case "J":
        case "N":
          range.values = values.toString().toUpperCase() as any;
          break;
        case "B":
        case "O":
          range.values = convertCase(values) as any;
          break;
        case "V":
          range.values = convertCase(values) as any;
          await context.sync();
          var driverRange = sheet.getRange("A4:V100");
          //sheet.protection.pauseProtection("CPM");
          var sortFields = [{ key: 0, ascending: true }];
          driverRange.sort.apply(sortFields);
          //sheet.protection.resumeProtection();
          break;
      }
      await context.sync();
    } catch (error) {
      console.log(error);
    }
  });
}

async function onStartTimeChange(event: Excel.WorksheetChangedEventArgs) {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    try {
      let address = event.address;
      if (event.triggerSource === "ThisLocalAddin") {
        return;
      }
      let sheet = context.workbook.worksheets.getItem("Schedule By Start Time");
      let details = event.details;
      let col = "";
      if (typeof details !== "undefined") {
        if (details.valueAfter === "STOP" || details.valueAfter === "stop" || details.valueAfter === "Stop") {
          let stopRange = sheet.getRange("A1");
          stopRange.load("values");
          await context.sync();
          let intervalID = stopRange.values;
          clearInterval(intervalID[0][0].valueOf());
          await context.sync();
          return;
        } else {
          col = address.substring(0, 1);
          var value = details.valueAfter;
        }
      } else {
        col = "E";
      }
      let inputRange = sheet.getRange(address);
      switch (col) {
        case "B":
          inputRange.values = convertCase(value) as any;
          break;
        case "D":
        case "F":
          inputRange.values = value.toString().toUpperCase() as any;
          break;
        case "E":
          updateStatus();
          break;
      }
      context.sync();
    } catch (error) {
      console.error(error);
    }
  });
}

async function onTUCSelection(event: Excel.WorksheetSelectionChangedEventArgs) {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    let address = event.address;
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    try {
      sheet.protection.pauseProtection("CPM");
      let range = sheet.getRange(address);
      range.load("values");
      await context.sync();
      let tempvalue = range.values;
      let value = tempvalue.toString();
      switch (address) {
        case "AB301":
        case "D301":
          if (value === "-") {
            sheet.getRange("D:AA").columnHidden = true;
            let groupRange = sheet.getRange("AB301");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("D:AA").columnHidden = false;
            let ungroupRange = sheet.getRange("AB301");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "AC301":
        case "BA301":
          if (value === "-") {
            sheet.getRange("AC:AZ").columnHidden = true;
            let groupRange = sheet.getRange("BA301");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("AC:AZ").columnHidden = false;
            let ungroupRange = sheet.getRange("BA301");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "BB301":
        case "BZ301":
          if (value === "-") {
            sheet.getRange("BB:BY").columnHidden = true;
            let groupRange = sheet.getRange("BZ301");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("BB:BY").columnHidden = false;
            let ungroupRange = sheet.getRange("BZ301");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "CA301":
        case "CY301":
          if (value === "-") {
            sheet.getRange("CA:CX").columnHidden = true;
            let groupRange = sheet.getRange("CY301");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("CA:CX").columnHidden = false;
            let ungroupRange = sheet.getRange("CY301");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "CZ301":
        case "DX301":
          if (value === "-") {
            sheet.getRange("CZ:DW").columnHidden = true;
            let groupRange = sheet.getRange("DX301");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("CZ:DW").columnHidden = false;
            let ungroupRange = sheet.getRange("DX301");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "DY301":
        case "EW301":
          if (value === "-") {
            sheet.getRange("DY:EV").columnHidden = true;
            let groupRange = sheet.getRange("EW301");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("DY:EV").columnHidden = false;
            let ungroupRange = sheet.getRange("EW301");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "EX301":
        case "FV301":
          if (value === "-") {
            sheet.getRange("EX:FU").columnHidden = true;
            let groupRange = sheet.getRange("FV301");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("EX:FU").columnHidden = false;
            let ungroupRange = sheet.getRange("FV301");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
      }
      await context.sync();
      sheet.protection.resumeProtection();
    } catch (error) {
      console.error(error);
      await context.sync();
      sheet.protection.resumeProtection();
    }
  });
}

async function onDUCSelection(event: Excel.WorksheetSelectionChangedEventArgs) {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    let address = event.address;
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    try {
      sheet.protection.pauseProtection("CPM");
      let range = sheet.getRange(address);
      range.load("values");
      await context.sync();
      let tempvalue = range.values;
      let value = tempvalue.toString();
      switch (address) {
        case "AE201":
        case "G201":
          if (value === "-") {
            sheet.getRange("G:AD").columnHidden = true;
            let groupRange = sheet.getRange("AE201");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("G:AD").columnHidden = false;
            let ungroupRange = sheet.getRange("AE201");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "AF201":
        case "BD201":
          if (value === "-") {
            sheet.getRange("AF:BC").columnHidden = true;
            let groupRange = sheet.getRange("BD201");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("AF:BC").columnHidden = false;
            let ungroupRange = sheet.getRange("BD201");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "BE201":
        case "CC201":
          if (value === "-") {
            sheet.getRange("BE:CB").columnHidden = true;
            let groupRange = sheet.getRange("CC201");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("BE:CB").columnHidden = false;
            let ungroupRange = sheet.getRange("CC201");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "CD201":
        case "DB201":
          if (value === "-") {
            sheet.getRange("CD:DA").columnHidden = true;
            let groupRange = sheet.getRange("DB201");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("CD:DA").columnHidden = false;
            let ungroupRange = sheet.getRange("DB201");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "DC201":
        case "EA201":
          if (value === "-") {
            sheet.getRange("DC:DZ").columnHidden = true;
            let groupRange = sheet.getRange("EA201");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("DC:DZ").columnHidden = false;
            let ungroupRange = sheet.getRange("EA201");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "EB201":
        case "EZ201":
          if (value === "-") {
            sheet.getRange("EB:EY").columnHidden = true;
            let groupRange = sheet.getRange("EZ201");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("EB:EY").columnHidden = false;
            let ungroupRange = sheet.getRange("EZ201");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
        case "FA201":
        case "FY201":
          if (value === "-") {
            sheet.getRange("FA:FX").columnHidden = true;
            let groupRange = sheet.getRange("FY201");
            groupRange.values = "+" as any;
            groupRange.getOffsetRange(1, 0).select();
          } else {
            sheet.getRange("FA:FX").columnHidden = false;
            let ungroupRange = sheet.getRange("FY201");
            ungroupRange.values = "-" as any;
            ungroupRange.getOffsetRange(1, 0).select();
          }
          break;
      }
      await context.sync();
      sheet.protection.resumeProtection();
    } catch (error) {
      console.error(error);
      await context.sync();
      sheet.protection.resumeProtection();
    }
  });
}

function convertCase(str) {
  var lower = String(str).toLowerCase();
  return lower.replace(/(^| )(\w)/g, function (x) {
    return x.toUpperCase();
  });
}

async function deleteRow() {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    let sheet = context.workbook.worksheets.getItem("Drivers");
    let driverRange = sheet.getRange("A4:P100");
    const sortFields = [{ key: 0, ascending: true }];
    driverRange.sort.apply(sortFields);
    await context.sync();
    let dataRange = sheet.getRange("$M$4");
    dataRange.autoFill("$M4:$M100");
    await context.sync();
  });
}

async function eventsOn() {
  Excel.run(async (context) => {
    context.runtime.load("enableEvents");
    await context.sync();
    context.runtime.enableEvents = true;
  });
}

async function digitalClock() {
  await Excel.run({ delayForCellEdit: true }, async (context) => {
    let sheet = context.workbook.worksheets.getItem("Schedule By Start Time");
    updateStatus();
    let range = sheet.getRange("F1");
    range.load("values");
    await context.sync();
    let cycles = range.values;
    let newcycles = +cycles + 5;
    range.values = newcycles as any;
    await context.sync();
  });
}

async function updateStatus() {
  Excel.run({ delayForCellEdit: true }, async (context) => {
    try {
      //let startTime = new Date();
      let sheet = context.workbook.worksheets.getItem("Schedule By Start Time");
      let driverCountRange = sheet.getRange("AW2");
      let statusCountRange = sheet.getRange("C101");
      driverCountRange.load("values");
      statusCountRange.load("values");
      await context.sync();
      let driversRow = driverCountRange.values;
      let statusRow = statusCountRange.values;
      if (+statusRow === 103) {
        statusRow = "104" as any;
      }
      let driverRange = sheet.getRange("AW3:BA" + driversRow);
      let statusRange = sheet.getRange("B104:F" + statusRow);
      driverRange.load("values");
      statusRange.load("values");
      await context.sync();
      let drivers = driverRange.values;
      let status = statusRange.values;
      let statusLength = status.length;
      let out1: string[][] = [];
      if (status[0][0] !== "") {
        out1 = status.concat(drivers);
      } else {
        out1 = drivers;
      }
      let out1Length = out1.length;
      let out2: string[][] = [];
      let duplicate = false;
      let currentTime = new Date();
      let currentMinutes = currentTime.getMinutes() + "";
      let currentHours = currentTime.getHours() + "";
      if (currentMinutes.length < 2) {
        currentMinutes = "0" + currentMinutes;
      }
      if (currentHours.length < 2) {
        currentHours = "0" + currentHours;
      }
      let timeString = currentHours + currentMinutes;
      for (let x = 0; x < out1Length; x++) {
        if (out1[x][0] === "" && out1[x][1] === "" && out1[x][2] === "" && out1[x][3] === "" && out1[x][4] === "") {
          continue;
        }
        let startPlus12 = +out1[x][1] + 1200;
        if (startPlus12 < 2400) {
          if (+timeString > startPlus12 && +timeString > +out1[x][3] && out1[x][1] !== "") {
            continue;
          }
        } else {
          let timeNumber = 0;
          if (+timeString < 1200) {
            timeNumber = +timeString + 2400;
          } else {
            timeNumber = +timeString;
          }
          let arriveTime = 0;
          if (+out1[x][3] < 1200) {
            arriveTime = +out1[x][3] + 2400;
          } else {
            arriveTime = +out1[x][3];
          }
          if (timeNumber > startPlus12 && timeNumber > arriveTime && out1[x][1] !== "") {
            continue;
          }
        }
        let a = out1[x][0].toString().toLowerCase();
        let out2Length = out2.length;
        for (let y = 0; y < out2Length; y++) {
          if (a === out2[y][0].toString().toLowerCase()) {
            duplicate = true;
            break;
          }
        }
        if (!duplicate) {
          out2.push([out1[x][0], out1[x][1], out1[x][2], out1[x][3], out1[x][4]]);
        } else {
          duplicate = false;
        }
      }
      out2.sort(function (x, y) {
        let xa = x[4];
        let ya = y[4];
        let result = xa == ya ? 0 : xa < ya ? -1 : 1;

        if (result == 0) {
          let xc = x[3];
          let yc = y[3];
          if (xc === "" && yc !== "") {
            result = 1;
          } else {
            if (xc !== "" && yc === "") {
              result = -1;
            } else {
              if (+xc - +yc > 1200) {
                result = -1;
              } else {
                if (+yc - +xc > 1200) {
                  result = 1;
                } else {
                  result = xc == yc ? 0 : xc < yc ? -1 : 1;
                }
              }
            }
          }
        }
        if (result == 0) {
          let xb = x[1];
          let yb = y[1];
          result = xb == yb ? 0 : xb < yb ? -1 : 1;
        }

        return result;
      });
      let out2Length = out2.length;
      if (statusLength > out2Length) {
        for (let x = out2Length; x < statusLength; x++) {
          out2.push(["", "", "", "", ""]);
        }
      }
      let outRange = sheet.getRange("B104:F" + (103 + out2.length));
      outRange.values = out2;
      await context.sync();
      //let endTime = new Date();
      //console.log(endTime.getTime() - startTime.getTime());
    } catch (error) {
      console.error(error);
    }
  });
}
