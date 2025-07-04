/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  // insertOrReplaceDataByHeader([]);
  // return
  try {
    await Excel.run(async (context) => {
      let ws = context.workbook.worksheets.getItem("DrugDetails");
      let packageDetails = context.workbook.worksheets.getItem("packageDistribution");
      let packageDetailsRange = packageDetails.getRange("A2:D7");
      let usedRange = ws.getUsedRange().getLastRow();
      let wsAutoReplenishHistroy = context.workbook.worksheets.getItem("AutoReplenishHistory");
      let drugsExpirationPredictions = context.workbook.worksheets.getItem(
        "Drug Replenish Dates(New Kits)"
      );
      let wsAutoReplenishMedGroups = context.workbook.worksheets.getItem(
        "auto_replenish_med_groups"
      );
      let wsRevenuePredictions = context.workbook.worksheets.getItem("Revenue Prediction");
      // let wsAutoReplenishMedGroupsAndPredictions = context.workbook.worksheets.getItem(
      //   "autoReplenish+Predictions"
      // );
      wsRevenuePredictions.getRangeByIndexes(1, 0, 10000, 50).clear(Excel.ClearApplyTo.contents);
      drugsExpirationPredictions
        .getRangeByIndexes(1, 0, 10000, 50)
        .clear(Excel.ClearApplyTo.contents);
      //Get the Details
      usedRange.load("rowIndex");
      await context.sync();
      let lastRow = usedRange.rowIndex;
      let data = ws.getRange(`B${1}:O${lastRow + 1}`);
      data.load("values");
      packageDetailsRange.load("values");
      await context.sync();
      let packageDetailsData = packageDetailsRange.values;
      let medsObj = {};
      let emkDetails = {};
      //Get the drug details
      console.log(data.values);
      data.values.forEach((row) => {
        medsObj[row[0]] = {
          totalUnitCost: row[3],
          laCarte: row[4],
          includedInPackages: [],
          shelfLife: row[7],
        };
        for (let i = 8; i <= 13; i++) {
          if (row[i].toString().trim() !== "") {
            medsObj[row[0]].includedInPackages.push(data.values[0][i]);
          }
        }
      });

      packageDetailsData.forEach((row) => {
        //Create the emk objecst
        emkDetails[row[0]] = {
          retailPrice: row[1],
          newKitShares: row[2],
          purchasePrice: row[3],
          drugs: [],
        };
      });
      console.log(medsObj,"Meds Object")

      //Get the New Kit Data
      let wsNewKit = context.workbook.worksheets.getItem("New Kit Data");
      let newKitsLastRow = wsNewKit.getUsedRange().getLastRow();
      newKitsLastRow.load("rowIndex");
      await context.sync();
      let newKitsLastRowIndex = newKitsLastRow.rowIndex;
      let dataRange = wsNewKit.getRange(`A2:B${newKitsLastRowIndex + 1}`);
      dataRange.load("values");
      await context.sync();
      let newKitData = dataRange.values;
      let salesHistory = {};
      //Get the Kit Revenue for each Kit and total Revenue
      let calculatedKitData = newKitData.map((row) => {
        salesHistory[formatDate(excelSerialDateToJSDate(row[0]))] = row[1];
        let numberOfKits = row[1];
        let EMK1 =
          Math.floor(emkDetails["EMK1"].newKitShares * numberOfKits) *
          emkDetails["EMK1"].retailPrice;
        let EMK5 =
          Math.floor(emkDetails["EMK5"].newKitShares * numberOfKits) *
          emkDetails["EMK5"].retailPrice;
        let EMK10 =
          Math.floor(emkDetails["EMK10"].newKitShares * numberOfKits) *
          emkDetails["EMK10"].retailPrice;
        let EMK15 =
          Math.floor(emkDetails["EMK15"].newKitShares * numberOfKits) *
          emkDetails["EMK15"].retailPrice;
        let EMK1Mini =
          Math.floor(emkDetails["EMK1-Mini"].newKitShares * numberOfKits) *
          emkDetails["EMK1-Mini"].retailPrice;
        let EMK10Mini =
          Math.floor(emkDetails["EMK10-Mini"].newKitShares * numberOfKits) *
          emkDetails["EMK10-Mini"].retailPrice;
        return [
          row[0],
          row[1],
          EMK1 + EMK5 + EMK10 + EMK15 + EMK1Mini + EMK10Mini,
          "",
          EMK1,
          EMK5,
          EMK10,
          EMK15,
          EMK1Mini,
          EMK10Mini,
        ];
      });
      //Add the Kit Revenue to the sheet

      wsNewKit.getRange("A2:J" + (calculatedKitData.length + 1)).values = calculatedKitData;
      //Add the total  Revenue to the sheet
      // const revenueLedger = calcRevenue(packages.emk1, salesHistory, projectedSales);
      // console.log(revenueLedger);
      //Get the drugs that belong to each Kit
      data.values.forEach((row) => {
        row[8] === "X" ? emkDetails["EMK1"]["drugs"].push(row[0]) : "";
        row[9] === "X" ? emkDetails["EMK5"]["drugs"].push(row[0]) : "";
        row[10] === "X" ? emkDetails["EMK10"]["drugs"].push(row[0]) : "";
        row[11] === "X" ? emkDetails["EMK15"]["drugs"].push(row[0]) : "";
        row[12] === "X" ? emkDetails["EMK1-Mini"]["drugs"].push(row[0]) : "";
        row[13] === "X" ? emkDetails["EMK10-Mini"]["drugs"].push(row[0]) : "";
      });
      //Creating calculation for all drugs per month
      let newKitDrugPredictions = [];
      Object.keys(salesHistory).forEach((month) => {
        let totalKitAmount = salesHistory[month];
        Object.keys(emkDetails).forEach((kit) => {
          let kitAmount = Math.floor(totalKitAmount * emkDetails[kit].newKitShares);
          if (kitAmount < 1) return;
          emkDetails[kit].drugs.forEach((drug) => {
            if (medsObj[drug].shelfLife == "" || medsObj[drug].shelfLife == "N/A") return;
            newKitDrugPredictions.push([
              month,
              kit,
              drug,
              kitAmount,
              medsObj[drug].laCarte * kitAmount,
              medsObj[drug].shelfLife,
            ]);
          });
        });
      });

      //Adding Replenish Dates to the Drug Details
      const updatedDrugData = newKitDrugPredictions.map((row) => {
        const [date, code, description, qty, total, expiryDays] = row;
        const [year, month] = date.split("-").map(Number);
        const baseDate = new Date(year, month - 1);

        const replenishments = [];

        for (let i = 1; i <= 10; i++) {
          const expireDate = new Date(baseDate);
          expireDate.setDate(expireDate.getDate() + expiryDays * i);

          const expireYear = expireDate.getFullYear();
          const expireMonth = String(expireDate.getMonth() + 1).padStart(2, "0");

          replenishments.push(`${expireYear}-${expireMonth}`);
        }

        return [...row, ...replenishments];
      });
      drugsExpirationPredictions.getRangeByIndexes(
        1,
        0,
        updatedDrugData.length,
        updatedDrugData[0].length
      ).values = updatedDrugData;
      //Get the history of auto replenishments
        // Get the used range of the worksheet
      // const autoReplenishmentsHistroyusedRange = wsAutoReplenishHistroy.getUsedRange();
      // usedRange.load("values");

      //   await context.sync();
      // const firstDataRow = usedRange.values[1];
      // let column = getAutoReplenishmentHistoryColumn(firstDataRow);
      // if(column !== -1){
      //   //Add it to the column number 
      // }
      // else{

      // }
      // --- Step 5: Execute everything
      const baseMap = getBaseKitMap(calculatedKitData);
      const forecastMap = generateForecast("2025-07", 300, baseMap);

      // Plug in your generated updatedDrugData (with replenishment dates)
      let drugDataMap = applyDrugDataRevenue(forecastMap, updatedDrugData);
      let usedRangeAutoReplenishMedGroups = wsAutoReplenishMedGroups.getUsedRange();
      let lastRowAutoReplenishMedGroups = usedRangeAutoReplenishMedGroups.getLastRow();
      lastRowAutoReplenishMedGroups.load("rowIndex");
      await context.sync();
      //Get Autor replenish sheet data
      let rangeAutoReplenishMedGroupsAll = wsAutoReplenishMedGroups.getRangeByIndexes(
        2,
        0,
        lastRowAutoReplenishMedGroups.rowIndex - 1,
        6
      );
      rangeAutoReplenishMedGroupsAll.load("values");
      await context.sync();
      //TODO
      //Add the Future expiration dates for the auto replenishments
      const outputAutoReplenishAndForecast = [
        [
          "Group",
          "Company",
          "Medication",
          "Expiration",
          "Price",
          "Auto Replenish",
          "Generated Dates",
        ],
      ];
      // wsAutoReplenishMedGroupsAndPredictions.getUsedRange().clear(Excel.ClearApplyTo.contents);
      for (const row of rangeAutoReplenishMedGroupsAll.values) {
        const [group, company, medication, expirationStr, price, autoReplenish] = row;

        if (autoReplenish !== "Enabled" || expirationStr == "N/A" || expirationStr == "") continue;

        const config = medsObj[medication];
        const baseDate = excelSerialDateToJSDate(expirationStr);

        // Always include original
        outputAutoReplenishAndForecast.push([
          group,
          company,
          medication,
          baseDate ? formatMonth(baseDate) : "N/A",
          price,
          autoReplenish,
          "",
        ]);

        if (!config || !baseDate){ 
      
          continue
        };

        const { shelfLife, laCarte: configPrice } = config;

        for (let i = 1; i <= 20; i++) {
          const futureDate = addDays(baseDate, shelfLife * i);
          outputAutoReplenishAndForecast.push([
            group,
            company,
            medication,
            formatMonth(futureDate),
            parseFloat(price.toFixed(2)),
            autoReplenish,
            "Generated",
          ]);
        }
      }


      // Auto-replenish items (only applied once)
      let autoReplenish = applyAutoReplenishOnce(forecastMap, outputAutoReplenishAndForecast);
      console.log(drugDataMap, baseMap, autoReplenish);
      // 1. Combine all unique months

      // --- Step 6: Final Output
      // const finalRevenueForecast = Array.from(forecastMap.entries()).map(([month, revenue]) => [month, revenue]);

      const allMonths = new Set([
        ...drugDataMap.keys(),
        ...autoReplenish.keys(),
        ...baseMap.keys(),
        ...forecastMap.keys(),
      ]);

      // 2. Generate final forecast array
      const finalRevenueForecast = [];

      for (const month of [...allMonths].sort()) {
        const newkit = baseMap.get(month) || 0;
        const auto = autoReplenish.get(month) || 0;
        const drugData = drugDataMap.get(month) || 0;
        const totalRevenue = newkit + auto + drugData;

        finalRevenueForecast.push([month, totalRevenue, newkit, auto, drugData]);
      }

      wsRevenuePredictions.getRangeByIndexes(
        1,
        0,
        finalRevenueForecast.length,
        finalRevenueForecast[0].length
      ).values = finalRevenueForecast;
      await context.sync();
      console.table(finalRevenueForecast);

      //AutoReplenish History
      let autoReplenishDatesAndData = getDatesAndData(finalRevenueForecast);
      const sheet = context.workbook.worksheets.getItem("AutoReplenishHistory");
    //Get last column
    const autoReplenishHistoryUsedRange = sheet.getUsedRange();
  autoReplenishHistoryUsedRange.load(["columnIndex", "columnCount","rowCount"]);
  await context.sync();
        // Load row 2 headers (row index 1)
    if(autoReplenishHistoryUsedRange.columnCount <=1) return //There is no data.
    const headerRange = sheet.getRangeByIndexes(1, 1, 1, autoReplenishHistoryUsedRange.columnCount-1);
    const datesRange = sheet.getRangeByIndexes(2,0,autoReplenishHistoryUsedRange.rowCount-2,1);
    headerRange.load("values");
    datesRange.load("values");
    await context.sync();
    console.log(datesRange.values)
    const headers = headerRange.values[0] // Row 2
    let currentMonth = formatDateWithDay(new Date());
    let currentMonthPostion = headers.indexOf(currentMonth) 
    let   datesInHistory = datesRange.values.map(x=>x[0]);
    let beginningDateIndex = datesInHistory.indexOf(autoReplenishDatesAndData[0][0]);
    console.log(currentMonth=headers[0])
    await context.sync();
    //Check if it exists in the headers
    if(currentMonthPostion !== -1){ //it Exists
      //Add data to the same column
      sheet.getRangeByIndexes(2+beginningDateIndex,currentMonthPostion+1,autoReplenishDatesAndData.length,1).values = autoReplenishDatesAndData.map(x=>[x[1]])
      sheet.getRangeByIndexes(2+beginningDateIndex,0,autoReplenishDatesAndData.length,1).values = autoReplenishDatesAndData.map(x=>[x[0]])
      
    }
    else{
      //Add data to a new column(The last column) starting at index 1 
      //Get the position in the dates column of the current starting date in the date array parameter
      //Add dates to it and add the data from the index of that row
      const currentMonthFormatted = formatDateWithDay(new Date());

    sheet.getRangeByIndexes(1, autoReplenishHistoryUsedRange.columnCount, 1, 1).numberFormat = [["@"]]; // force text
    sheet.getRangeByIndexes(1, autoReplenishHistoryUsedRange.columnCount, 1, 1).values = [[currentMonthFormatted]];
    sheet.getRangeByIndexes(2+beginningDateIndex,autoReplenishHistoryUsedRange.columnCount,autoReplenishDatesAndData.length,1).values = autoReplenishDatesAndData.map(x=>[x[1]])
    sheet.getRangeByIndexes(2+beginningDateIndex,0,autoReplenishDatesAndData.length,1).values = autoReplenishDatesAndData.map(x=>[x[0]])
    }
        
    await context.sync();

      
      // const BATCH_SIZE = 10000;

      // for (let startRow = 0; startRow < outputAutoReplenishAndForecast.length; startRow += BATCH_SIZE) {
      //     const chunk = outputAutoReplenishAndForecast.slice(startRow, startRow + BATCH_SIZE);
          
      //     wsAutoReplenishMedGroupsAndPredictions
      //         .getRangeByIndexes(startRow, 0, chunk.length, chunk[0].length)
      //         .values = chunk;
      
      //     await context.sync();
      // }
      
      return context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
function getDatesAndData(foreCastData=[[]]){
  return foreCastData.map(
    x => [x[0],x[3]]
  )
}
function getBaseKitMap(baseKitRevenue) {
  const map = new Map();

  const now = new Date();
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1); // First day of next month

  baseKitRevenue.forEach(([dateStr, kitQuantity, revenue]) => {
    const date = excelSerialDateToJSDate(dateStr);
    console.log("Here we are")
    if (date >= nextMonth) {
      const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
      map.set(key, revenue);
    }
  });

  return map;
}


// --- Step 2: Forecast structure (June 2023 → May 2033)
function generateForecast(start = "2023-06", months = 120, baseMap = new Map()) {
  const forecast = new Map();
  const [startYear, startMonth] = start.split("-").map(Number);
  const date = new Date(startYear, startMonth - 1);

  for (let i = 0; i < months; i++) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const key = `${year}-${month}`;
    const baseRevenue = baseMap.get(key) || 0;
    forecast.set(key, baseRevenue);
    date.setMonth(date.getMonth() + 1);
  }

  return forecast;
}

// --- Step 3: Add drugData replenishment costs
function applyDrugDataRevenue(forecastMap, drugData) {
  let drugDataMap = new Map();
  for (const row of drugData) {
    const total = parseFloat(row[4]);
    const replenishmentDates = row.slice(6);
    // dynamically added dates
    replenishmentDates.forEach((date) => {
      if (forecastMap.has(date)) {
        forecastMap.set(date, forecastMap.get(date) + total);
        drugDataMap.set(
          date,
          drugDataMap.get(date) != undefined ? drugDataMap.get(date) + total : total
        );
      }
    });
  }
  return drugDataMap;
}

// --- Step 4: Add Auto Replenish (just once, at expiration date)
function applyAutoReplenishOnce(forecastMap, autoData) {
  let autoReplenish = new Map();
  autoData.forEach((row) => {
    const [Group, Company, Medication, expDate, priceStr, status] = row;

    if (status !== "Enabled") return;
    const price = typeof priceStr == "string" ? parseFloat(priceStr.replace("$", "")) : priceStr;

    // const [expMonth, , expYear] = expDate.split("/").map(Number);
    // const key = `${expYear}-${String(expMonth).padStart(2, '0')}`;
    const date = new parseMonth(expDate);
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (forecastMap.has(key)) {
      if (!isNaN(price)) {
        forecastMap.set(key, forecastMap.get(key) + price);
        autoReplenish.set(
          key,
          autoReplenish.get(key) !== undefined ? autoReplenish.get(key) + price : price
        );
      }
    }
  });
  return autoReplenish;
}

// // --- Step 5: Execute everything
// const baseMap = getBaseKitMap(baseKitRevenue);
// const forecastMap = generateForecast("2023-06", 120, baseMap);

// // Plug in your generated updatedDrugData (with replenishment dates)
// applyDrugDataRevenue(forecastMap, updatedDrugData);

// // Auto-replenish items (only applied once)
// applyAutoReplenishOnce(forecastMap, [
//   ["42", "Dental Depot", "Insta-Glucose", "2/28/2026", "$10.85", "Enabled"],
//   ["42", "Dental Depot", "Nitroglycerin Sublingual Tablets 0.4 mg", "5/31/2026", "$46.71", "Enabled"],
//   ["42", "Dental Depot", "Albuterol Sulfate (60 doses)", "5/31/2026", "$79.61", "Enabled"],
//   ["42", "Dental Depot", "Ammonia Towelette", "3/31/2027", "$14.08", "Enabled"],
//   ["42", "Dental Depot", "Adrenaline 1 mg/mL", "6/30/2026", "$31.27", "Enabled"],
//   ["42", "Dental Depot", "Adrenaline 1 mg/mL", "6/30/2026", "$31.27", "Enabled"],
//   ["42", "Dental Depot", "Naloxone HCL 0.4 mg/mL", "4/30/2026", "$43.45", "Enabled"],
// ]);

// // --- Step 6: Final Output
// const finalRevenueForecast = Array.from(forecastMap.entries()).map(([month, revenue]) => [month, revenue]);
// console.table(finalRevenueForecast);

// ─── Helpers ────────────────────────────────────────────────────────────────
function parseMonth(ym) {
  // More robust parsing of YYYY-MM strings
  const [y, m] = ym.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, 1));
}
// Instead of using new Date() directly:
function getCurrentMonthUTC() {
  const now = new Date();
  return new Date(Date.UTC(now.getFullYear(), now.getMonth(), 1));
}
function formatMonth(dt) {
  const y = dt.getUTCFullYear(),
    m = String(dt.getUTCMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}
function addDays(dt, n) {
  return new Date(dt.valueOf() + n * 864e5);
}
function addMonths(dt, n) {
  const y = dt.getUTCFullYear(),
    mo = dt.getUTCMonth() + n;
  return new Date(Date.UTC(y + Math.floor(mo / 12), mo % 12, 1));
}
function generateProjections(start, end, perMonth) {
  const result = {};
  let cur = parseMonth(start),
    last = parseMonth(end);
  while (cur <= last) {
    result[formatMonth(cur)] = perMonth;
    cur = addMonths(cur, 1);
  }
  return result;
}
function formatDate(date) {
  // Always use UTC to avoid timezone issues
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}
function formatDateWithDay(date) {
  // Always use UTC to avoid timezone issues
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = date.getUTCDate()
  return `${year}-${month}-${day}`;
}
function excelSerialDateToJSDate(serial) {
  // UTC-based conversion to avoid timezone issues
  const utcDays = Math.floor(serial - 25569); // 25569 = days between 1900 and 1970
  const utcValue = utcDays * 86400; // 86400 = seconds per day
  const dateInfo = new Date(utcValue * 1000);
  
  // Create a new date using UTC values to avoid timezone offset
  return new Date(Date.UTC(
      dateInfo.getUTCFullYear(),
      dateInfo.getUTCMonth(),
      dateInfo.getUTCDate()
  ));
}
/**
 * @param {Array} headers - List of dates structures as month-year
 * 
 */
function getAutoReplenishmentHistoryColumn(headers){
  let currentDate = new Date();
  let dateStr= formatDate(currentDate);
  //Check the headers of the sheet to see if the date already exists

  let index =  headers.indexOf(dateStr);
  return index
}
// ─── run it ────────────────────────────────────────────────────────────────
async function insertOrReplaceDataByHeader(dates, Data =[300,800]) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("AutoReplenishHistory");
    //Get last column
    const usedRange = sheet.getUsedRange();
  usedRange.load(["columnIndex", "columnCount"]);
  await context.sync();
        // Load row 2 headers (row index 1)
        if(usedRange.columnCount <=1) return //There is no data.
    const headerRange = sheet.getRangeByIndexes(1, 1, 1, usedRange.columnCount-1);
    headerRange.load("values");
    await context.sync();
    console.log(headerRange.values)
    const headers = headerRange.values[0] // Row 2
    let currentMonth = formatDate(getCurrentMonthUTC());
    let currentMonthPostion = headers.indexOf(currentMonth) 
    console.log(currentMonth=headers[0])
    //Check if it exists in the headers
    if(currentMonthPostion !== -1){ //it Exists
      //Add data to the same column
      console.log(Data.map(x=>[x]))
      sheet.getRangeByIndexes(2,currentMonthPostion+1,Data.length,1).values = Data.map(x=>[x])
    }
    else{
      //Add data to a new column(The last column) starting at index 1 
      //Get the position in the dates column of the current starting date in the date array parameter
      //Add dates to it and add the data from the index of that row
    }
        await context.sync();
  });
}
