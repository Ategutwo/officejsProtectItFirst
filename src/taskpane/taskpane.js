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
        let currentMonthKey = formatDate(excelSerialDateToJSDate(row[0]))
        if(!isPastMonth(currentMonthKey)){
          salesHistory[currentMonthKey] = row[1];
        }
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
        const baseDate = new Date(Date.UTC(year, month - 1, 1)); // Use UTC

        const replenishments = [];

        for (let i = 1; i <= 10; i++) {
          const expireDate = new Date(baseDate);
          expireDate.setUTCDate(expireDate.getUTCDate() + expiryDays * i); // Use UTC

          const expireYear = expireDate.getUTCFullYear();
          const expireMonth = String(expireDate.getUTCMonth() + 1).padStart(2, "0");

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
      const allMonths = new Set([
        ...drugDataMap.keys(),
        ...autoReplenish.keys(),
        ...baseMap.keys(),
        ...forecastMap.keys(),
      ]);

      // 2. Generate final forecast array with Non-Auto Replenishment column
      const finalRevenueForecast = [];
      const currentMonth = getCurrentMonthUTC();
      const currentMonthKey = formatMonth(currentMonth);
      // Get the percentage from Non-Auto Replenish sheet
      let nonAutoReplenishSheet;
      let nonAutoPercentage = 0.25; // Default to 25%
      
      try {
        nonAutoReplenishSheet = context.workbook.worksheets.getItem("Non-Auto Replenish");
        const g2Range = nonAutoReplenishSheet.getRange("G2");
        const h2Range = nonAutoReplenishSheet.getRange("H2");
        
        g2Range.load("values");
        h2Range.load("values");
        await context.sync();
        
        // Use H2 if not empty, otherwise use G2
        if (h2Range.values[0][0] !== null && h2Range.values[0][0] !== "" && !isNaN(h2Range.values[0][0])) {
          nonAutoPercentage = parseFloat(h2Range.values[0][0]);
        } else if (g2Range.values[0][0] !== null && g2Range.values[0][0] !== "" && !isNaN(g2Range.values[0][0])) {
          nonAutoPercentage = parseFloat(g2Range.values[0][0]);
        }
        
        console.log("Using Non-Auto Replenishment percentage:", nonAutoPercentage);
      } catch (error) {
        console.log("Non-Auto Replenish sheet not found or error reading values, using default 25%");
      }

      // Get AED sales value
      let aedSalesValue = 0;
      try {
        const aedSalesSheet = context.workbook.worksheets.getItem("AED Sales");
        const aedSalesRange = aedSalesSheet.getRange("G2");
        aedSalesRange.load("values");
        await context.sync();
        
        if (aedSalesRange.values[0][0] !== null && aedSalesRange.values[0][0] !== "" && !isNaN(aedSalesRange.values[0][0])) {
          aedSalesValue = parseFloat(aedSalesRange.values[0][0]);
        }
        
        console.log("Using AED Sales value:", aedSalesValue);
      } catch (error) {
        console.log("AED Sales sheet not found or error reading G2 value, using 0");
      }

      for (const month of [...allMonths].sort()) {
        const newkit = baseMap.get(month) || 0;
        const auto = autoReplenish.get(month) || 0;
        const drugData = drugDataMap.get(month) || 0;
        
        // Calculate Non-Auto Replenishment (25% of Auto Replenish, only for future months)
        const nonAutoReplenishment = (month > currentMonthKey) ? auto * nonAutoPercentage : 0;
        
        // Calculate AED Sales (only for future months)
        const aedSales = (month > currentMonthKey) ? aedSalesValue : 0;
        
        // Update total revenue to include non-auto replenishment and AED sales
        const totalRevenue = newkit + auto + drugData + nonAutoReplenishment + aedSales;

        finalRevenueForecast.push([
          month, 
          totalRevenue, 
          newkit, 
          auto, 
          drugData,
          nonAutoReplenishment,  // New column: Non-Auto Replenishment
          aedSales               // New column: AED sales
        ]);
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
      
      console.log("Dates in History:", datesRange.values);
      const headers = headerRange.values[0] // Row 2
      let currentMonthFormatted = formatDateWithDay(new Date());
      let currentMonthPostion = headers.indexOf(currentMonthFormatted);
      let datesInHistory = datesRange.values.map(x => x[0]);
      
      // Use normalized date comparison to find the correct starting row
      let beginningDateIndex = findDateIndexInHistory(datesInHistory, autoReplenishDatesAndData[0][0]);
      
      console.log("Beginning Date Index:", beginningDateIndex);
      console.log("Target Date:", autoReplenishDatesAndData[0][0]);
      console.log("Current Month Position:", currentMonthPostion);
      
      await context.sync();
      
      //Check if it exists in the headers
      if(currentMonthPostion !== -1){ //it Exists
        //Add data to the same column
        sheet.getRangeByIndexes(2+beginningDateIndex,currentMonthPostion+1,autoReplenishDatesAndData.length,1).values = autoReplenishDatesAndData.map(x=>[x[1]])
        sheet.getRangeByIndexes(2+beginningDateIndex,0,autoReplenishDatesAndData.length,1).values = autoReplenishDatesAndData.map(x=>[x[0]])
      }
      else{
        //Add data to a new column(The last column) starting at index 1 
        
        sheet.getRangeByIndexes(1, autoReplenishHistoryUsedRange.columnCount, 1, 1).numberFormat = [["@"]]; // force text
        sheet.getRangeByIndexes(1, autoReplenishHistoryUsedRange.columnCount, 1, 1).values = [[currentMonthFormatted]];
        sheet.getRangeByIndexes(2+beginningDateIndex,autoReplenishHistoryUsedRange.columnCount,autoReplenishDatesAndData.length,1).values = autoReplenishDatesAndData.map(x=>[x[1]])
        sheet.getRangeByIndexes(2+beginningDateIndex,0,autoReplenishDatesAndData.length,1).values = autoReplenishDatesAndData.map(x=>[x[0]])
      }
      
      await context.sync();
      return context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}



// ... rest of the helper functions remain exactly the same ...

function getDatesAndData(foreCastData=[[]]){
  return foreCastData.map(
    x => [x[0],x[3]]
  )
}

function getBaseKitMap(baseKitRevenue) {
  const map = new Map();

  const now = getCurrentMonthUTC(); // Use UTC for consistency
  const nextMonth = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 1));

  baseKitRevenue.forEach(([dateStr, kitQuantity, revenue]) => {
    const date = excelSerialDateToJSDate(dateStr);
    if (date >= nextMonth) {
      const key = `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
      map.set(key, revenue);
    }
  });

  return map;
}

function generateForecast(start = "2023-06", months = 120, baseMap = new Map()) {
  const forecast = new Map();
  const [startYear, startMonth] = start.split("-").map(Number);
  const date = new Date(Date.UTC(startYear, startMonth - 1, 1)); // Use UTC

  for (let i = 0; i < months; i++) {
    const year = date.getUTCFullYear();
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    const key = `${year}-${month}`;
    const baseRevenue = baseMap.get(key) || 0;
    forecast.set(key, baseRevenue);
    date.setUTCMonth(date.getUTCMonth() + 1);
  }

  return forecast;
}

function applyDrugDataRevenue(forecastMap, drugData) {
  let drugDataMap = new Map();
  for (const row of drugData) {
    const total = parseFloat(row[4]);
    const replenishmentDates = row.slice(6);
    
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

function applyAutoReplenishOnce(forecastMap, autoData) {
  let autoReplenish = new Map();
  autoData.forEach((row) => {
    const [Group, Company, Medication, expDate, priceStr, status] = row;

    if (status !== "Enabled") return;
    const price = typeof priceStr == "string" ? parseFloat(priceStr.replace("$", "")) : priceStr;

    // Parse expiration date consistently
    const date = parseMonthString(expDate);
    if (!date) return;
    
    const key = `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}`;
    
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

// ─── Helper Functions (Timezone Safe) ──────────────────────────────────────

function excelSerialDateToJSDate(serial) {
  // More robust Excel serial date conversion with UTC
  if (!serial || isNaN(serial)) return new Date(NaN);
  
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  
  // Return UTC date to avoid timezone shifts
  return new Date(Date.UTC(
    dateInfo.getUTCFullYear(),
    dateInfo.getUTCMonth(),
    dateInfo.getUTCDate()
  ));
}

function isPastMonth(inputDate) {
  const today = getCurrentMonthUTC();
  
  const currentYear = today.getUTCFullYear();
  const currentMonth = today.getUTCMonth() + 1;

  let year, month;

  if (inputDate instanceof Date) {
    year = inputDate.getUTCFullYear();
    month = inputDate.getUTCMonth() + 1;
  } else if (typeof inputDate === "string") {
    [year, month] = inputDate.split("-").map(Number);
  } else {
    return false;
  }

  return (currentYear > year || (currentYear === year && currentMonth > month));
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
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatMonth(dt) {
  const y = dt.getUTCFullYear(),
    m = String(dt.getUTCMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

function addDays(dt, n) {
  // Use UTC to avoid DST issues
  const result = new Date(dt);
  result.setUTCDate(result.getUTCDate() + n);
  return result;
}

function getCurrentMonthUTC() {
  const now = new Date();
  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1));
}

function parseMonthString(ym) {
  // Robust parsing of various date formats
  if (!ym || ym === "N/A") return new Date(NaN);
  
  try {
    // Handle "YYYY-MM" format
    if (/^\d{4}-\d{2}$/.test(ym)) {
      const [y, m] = ym.split("-").map(Number);
      return new Date(Date.UTC(y, m - 1, 1));
    }
    
    // Handle "MM/DD/YYYY" or other formats
    const parsed = new Date(ym);
    if (isNaN(parsed.getTime())) return new Date(NaN);
    
    // Convert to UTC first day of month
    return new Date(Date.UTC(parsed.getUTCFullYear(), parsed.getUTCMonth(), 1));
  } catch (e) {
    return new Date(NaN);
  }
}

function findDateIndexInHistory(datesInHistory, targetDate) {
  // Normalize both dates for comparison
  const normalizedTarget = normalizeDateForComparison(targetDate);
  
  for (let i = 0; i < datesInHistory.length; i++) {
    const normalizedHistory = normalizeDateForComparison(datesInHistory[i]);
    if (normalizedTarget.getTime() === normalizedHistory.getTime()) {
      return i;
    }
  }
  
  console.warn("Target date not found in history, using index 0");
  return 0;
}

function normalizeDateForComparison(dateStr) {
  // Handle both "YYYY-MM" and "YYYY-MM-DD" formats
  const parts = dateStr.split('-');
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]);
  
  // Always return the first day of month in UTC for consistent comparison
  return new Date(Date.UTC(year, month - 1, 1));
}

// Legacy function kept for compatibility
function parseMonth(ym) {
  return parseMonthString(ym);
}