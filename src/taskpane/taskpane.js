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
      
      // Get AED Details sheet data
      let wsAEDDetails = context.workbook.worksheets.getItem("AED_Details(main)");
      let aedUsedRange = wsAEDDetails.getUsedRange();
      aedUsedRange.load(["rowIndex", "rowCount"]);
      await context.sync();
      let aedLastRow = aedUsedRange.rowIndex;
      let aedDataRange = wsAEDDetails.getRange(`A2:I${aedLastRow + 1}`);
      aedDataRange.load("values");
      await context.sync();
      
      // Create AED details object - use the AED Accessories name as key
      let aedMedsObj = {};
      aedDataRange.values.forEach((row) => {
        if (row[1] && row[1].toString().trim() !== "") {
          aedMedsObj[row[1].trim()] = {
            itemName: row[1],
            units: row[2],
            unitCost: row[3],
            totalUnitCost: row[4],
            laCarte: row[5],
            shelfLife: row[8],
            basePrice: row[5] // Store base price for monthly increases
          };
        }
      });
      
      console.log("AED Items loaded:", Object.keys(aedMedsObj));
      
      // Get monthly increase data from "Yearly Med Increase" sheet
      let monthlyIncreaseMap = new Map();
      try {
        const yearlyIncreaseSheet = context.workbook.worksheets.getItem("Yearly Med Increase");
        const usedRangeIncrease = yearlyIncreaseSheet.getUsedRange();
        usedRangeIncrease.load("rowCount");
        await context.sync();
        
        if (usedRangeIncrease.rowCount > 1) {
          const increaseDataRange = yearlyIncreaseSheet.getRange(`A2:B${usedRangeIncrease.rowCount}`);
          increaseDataRange.load("values");
          await context.sync();
          
          increaseDataRange.values.forEach(row => {
            if (row[0] && row[1] !== null && row[1] !== "" && !isNaN(row[1])) {
              const date = excelSerialDateToJSDate(row[0]);
              const monthKey = formatDate(date);
              //Percentage Increase Area
              monthlyIncreaseMap.set(monthKey, parseFloat(row[1]));
            }
          });
          console.log("Monthly increase data loaded:", monthlyIncreaseMap);
        }
      } catch (error) {
        console.log("Yearly Med Increase sheet not found or error reading data, using no increases");
      }
      
      // Clear only up to column J to preserve discount in column L
      wsRevenuePredictions.getRangeByIndexes(1, 0, 10000, 15).clear(Excel.ClearApplyTo.contents);
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
          basePrice: row[4] // Store base price for monthly increases
        };
        for (let i = 8; i <= 13; i++) {
          if (row[i] && row[i].toString().trim() !== "") {
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
          baseRetailPrice: row[1] // Store base price for monthly increases
        };
      });

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
      
      // Get discount percentage from L2
      let discountPercentage = 0;
      try {
        const discountRange = wsNewKit.getRange("L2");
        discountRange.load("values");
        await context.sync();
        
        if (discountRange.values[0][0] !== null && discountRange.values[0][0] !== "" && !isNaN(discountRange.values[0][0])) {
          discountPercentage = parseFloat(discountRange.values[0][0]);
          console.log("Using discount percentage:", discountPercentage);
        }
      } catch (error) {
        console.log("Error reading discount from L2, using 0% discount");
      }
      
      // Function to apply monthly price increases cumulatively
      const getAdjustedPrice = (basePrice, targetDate, increaseMap) => {
        if (increaseMap.size === 0) return basePrice;
        
        const currentDate = new Date();
        const targetMonthKey = formatDate(targetDate);
        
        // If target date is before current date, no increases applied
        if (targetDate < currentDate) return basePrice;
        
        // Sort all increase months chronologically
        const sortedIncreaseMonths = Array.from(increaseMap.keys()).sort();
        
        let adjustedPrice = basePrice;
        let cumulativeIncrease = 1;
        
        // Apply all increases that occur before or in the target month
        for (const increaseMonth of sortedIncreaseMonths) {
          if (increaseMonth <= targetMonthKey) {
            const increaseRate = increaseMap.get(increaseMonth);
            cumulativeIncrease *= (1 + increaseRate);
          }
        }
        
        adjustedPrice = basePrice * cumulativeIncrease;
        return adjustedPrice;
      };
      
      //Get the Kit Revenue for each Kit and total Revenue with discount applied
      let calculatedKitData = newKitData.map((row) => {
        let currentMonthKey = formatDate(excelSerialDateToJSDate(row[0]))
        if(!isPastMonth(currentMonthKey)){
          salesHistory[currentMonthKey] = row[1];
        }
        let numberOfKits = row[1];
        const rowDate = excelSerialDateToJSDate(row[0]);
        
        // Calculate base prices with monthly increases applied
        let EMK1_base_retail = getAdjustedPrice(emkDetails["EMK1"].baseRetailPrice, rowDate, monthlyIncreaseMap);
        let EMK5_base_retail = getAdjustedPrice(emkDetails["EMK5"].baseRetailPrice, rowDate, monthlyIncreaseMap);
        let EMK10_base_retail = getAdjustedPrice(emkDetails["EMK10"].baseRetailPrice, rowDate, monthlyIncreaseMap);
        let EMK15_base_retail = getAdjustedPrice(emkDetails["EMK15"].baseRetailPrice, rowDate, monthlyIncreaseMap);
        let EMK1Mini_base_retail = getAdjustedPrice(emkDetails["EMK1-Mini"].baseRetailPrice, rowDate, monthlyIncreaseMap);
        let EMK10Mini_base_retail = getAdjustedPrice(emkDetails["EMK10-Mini"].baseRetailPrice, rowDate, monthlyIncreaseMap);
        
        let EMK1_base =
          Math.floor(emkDetails["EMK1"].newKitShares * numberOfKits) *
          EMK1_base_retail;
        let EMK5_base =
          Math.floor(emkDetails["EMK5"].newKitShares * numberOfKits) *
          EMK5_base_retail;
        let EMK10_base =
          Math.floor(emkDetails["EMK10"].newKitShares * numberOfKits) *
          EMK10_base_retail;
        let EMK15_base =
          Math.floor(emkDetails["EMK15"].newKitShares * numberOfKits) *
          EMK15_base_retail;
        let EMK1Mini_base =
          Math.floor(emkDetails["EMK1-Mini"].newKitShares * numberOfKits) *
          EMK1Mini_base_retail;
        let EMK10Mini_base =
          Math.floor(emkDetails["EMK10-Mini"].newKitShares * numberOfKits) *
          EMK10Mini_base_retail;
        
        // Apply discount to total revenue
        let totalBaseRevenue = EMK1_base + EMK5_base + EMK10_base + EMK15_base + EMK1Mini_base + EMK10Mini_base;
        let totalDiscountedRevenue = totalBaseRevenue * (1 - discountPercentage);
        
        // Apply discount to individual kit revenues
        let EMK1 = EMK1_base * (1 - discountPercentage);
        let EMK5 = EMK5_base * (1 - discountPercentage);
        let EMK10 = EMK10_base * (1 - discountPercentage);
        let EMK15 = EMK15_base * (1 - discountPercentage);
        let EMK1Mini = EMK1Mini_base * (1 - discountPercentage);
        let EMK10Mini = EMK10Mini_base * (1 - discountPercentage);
        
        return [
          row[0],
          row[1],
          totalDiscountedRevenue,
          "",
          EMK1,
          EMK5,
          EMK10,
          EMK15,
          EMK1Mini,
          EMK10Mini,
        ];
      });
      
      //Add the Kit Revenue to the sheet (only up to column J to preserve discount in L)
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
// ========== NEW TAKEOVER KIT CALCULATIONS ==========
let takeoverRevenueMap = new Map();
let takeoverDrugPredictions = [];
 let drugsTakeoverExpirationPredictions = context.workbook.worksheets.getItem(
        "Replenish Dates(TakoverKits)"
      );
       drugsTakeoverExpirationPredictions
        .getRangeByIndexes(1, 0, 10000, 50)
        .clear(Excel.ClearApplyTo.contents);
        await context.sync();
try {
  // Get Takeover Drug Details
  let wsTakeoverDrugDetails = context.workbook.worksheets.getItem("TakeOverDrugDetails");
  let takeoverDrugsUsedRange = wsTakeoverDrugDetails.getUsedRange();
  takeoverDrugsUsedRange.load(["rowIndex", "rowCount"]);
  await context.sync();
  
  let takeoverDrugsLastRow = takeoverDrugsUsedRange.rowCount;
  let takeoverDrugsDataRange = wsTakeoverDrugDetails.getRange(`A2:C${takeoverDrugsLastRow + 1}`);
  takeoverDrugsDataRange.load("values");
  await context.sync();
  
  console.log(`Takeover Drug Details rows: ${takeoverDrugsDataRange.values.length}`);
  console.log("Takeover Drug Details sample:", takeoverDrugsDataRange.values.slice(0, 5));

  // Create takeover drug details object grouped by kit
  let takeoverDrugDetails = {};
  let drugCount = 0;
  
  takeoverDrugsDataRange.values.forEach((row, index) => {
    if (row[0] && row[0].toString().trim() !== "") {
      const drugName = row[0].toString().trim();
      const kit = row[1] ? row[1].toString().trim() : "";
      const daysToFirstReplenishment = row[2] && !isNaN(row[2]) ? parseInt(row[2]) : 0;
      
      if (kit) {
        if (!takeoverDrugDetails[kit]) {
          takeoverDrugDetails[kit] = [];
        }
        
        // Get drug price and shelf life from main drug details
        const drugPrice = medsObj[drugName] ? medsObj[drugName].laCarte : 0;
        const shelfLife = medsObj[drugName] ? medsObj[drugName].shelfLife : 0;
        
        if (drugPrice > 0 && daysToFirstReplenishment > 0) {
          takeoverDrugDetails[kit].push({
            drugName: drugName,
            daysToFirstReplenishment: daysToFirstReplenishment,
            price: drugPrice,
            shelfLife: shelfLife
          });
          drugCount++;
        } else {
          console.log(`Skipping drug ${drugName} - price: ${drugPrice}, days: ${daysToFirstReplenishment}`);
        }
      }
    }
  });

  console.log(`Loaded ${drugCount} valid takeover drugs across kits:`, Object.keys(takeoverDrugDetails));
  Object.keys(takeoverDrugDetails).forEach(kit => {
    console.log(`  ${kit}: ${takeoverDrugDetails[kit].length} drugs`);
  });

  // Get New Takeover Kit Data
  let wsNewTakeoverKitData = context.workbook.worksheets.getItem("New Takover Kit Data");
  let takeoverKitsUsedRange = wsNewTakeoverKitData.getUsedRange();
  takeoverKitsUsedRange.load(["rowIndex", "rowCount", "columnCount"]);
  await context.sync();
  
  let takeoverKitsLastRow = takeoverKitsUsedRange.rowCount;
  let takeoverKitsDataRange = wsNewTakeoverKitData.getRange(`A2:G${takeoverKitsLastRow + 1}`);
  takeoverKitsDataRange.load("values");
  await context.sync();
  
  console.log(`Takeover Kit Data rows: ${takeoverKitsDataRange.values.length}`);
  console.log("Takeover Kit Data sample:", takeoverKitsDataRange.values.slice(0, 3));

  // Process takeover kit data
  
  const kitColumns = {
    1: "EMK1",
    2: "EMK5", 
    3: "EMK10",
    4: "EMK15",
    5: "EMK1-Mini",
    6: "EMK10-Mini"
  };

  let totalProcessed = 0;
  let monthlyBreakdown = {};

  takeoverKitsDataRange.values.forEach((row, rowIndex) => {
    const takeoverDate = excelSerialDateToJSDate(row[0]);
    const takeoverMonthKey = formatDate(takeoverDate);
    
    console.log(`\nProcessing month: ${takeoverMonthKey}`);
    monthlyBreakdown[takeoverMonthKey] = { kits: 0, drugs: 0, revenue: 0 };
    
    // Process each kit type in this month
    for (let i = 1; i <= 6; i++) {
      const kitName = kitColumns[i];
      const numberOfTakeovers = row[i] && !isNaN(row[i]) ? parseInt(row[i]) : 0;
      
      console.log(`  ${kitName}: ${numberOfTakeovers} takeovers`);
      
      if (numberOfTakeovers > 0 && takeoverDrugDetails[kitName]) {
        monthlyBreakdown[takeoverMonthKey].kits += numberOfTakeovers;
        
        // For each drug in this kit, calculate replenishments
        takeoverDrugDetails[kitName].forEach(drug => {
          if (drug.price && !isNaN(drug.price) && drug.daysToFirstReplenishment > 0 && drug.shelfLife && !isNaN(drug.shelfLife)) {
            const drugRevenue = drug.price * numberOfTakeovers;
            totalProcessed++;
            monthlyBreakdown[takeoverMonthKey].drugs++;
            monthlyBreakdown[takeoverMonthKey].revenue += drugRevenue;
            
            // Calculate first replenishment date
            const firstReplenishDate = new Date(takeoverDate);
            firstReplenishDate.setUTCDate(firstReplenishDate.getUTCDate() + drug.daysToFirstReplenishment);
            const firstReplenishMonthKey = formatDate(firstReplenishDate);
            
            console.log(`    ${drug.drugName}: ${numberOfTakeovers} × $${drug.price} = $${drugRevenue}, first replenish: ${firstReplenishMonthKey}`);
            
            // Generate ALL replenishment dates
            const replenishmentDates = [];
            
            for (let i = 0; i < 10; i++) {
              const expireDate = new Date(firstReplenishDate);
              expireDate.setUTCDate(expireDate.getUTCDate() + drug.shelfLife * i);
              
              const expireYear = expireDate.getUTCFullYear();
              const expireMonth = String(expireDate.getUTCMonth() + 1).padStart(2, "0");
              const monthKey = `${expireYear}-${expireMonth}`;
              
              replenishmentDates.push(monthKey);
              
              // Add revenue to EACH replenishment month
              const currentRevenue = takeoverRevenueMap.get(monthKey) || 0;
              takeoverRevenueMap.set(monthKey, currentRevenue + drugRevenue);
            }
            
            // Add to drug predictions for tracking
            takeoverDrugPredictions.push([
              takeoverMonthKey,
              kitName + " (Takeover)",
              drug.drugName,
              numberOfTakeovers,
              drugRevenue,
              drug.shelfLife,
              drug.daysToFirstReplenishment,
              ...replenishmentDates
            ]);
          }
        });
      }
    }
  });

  console.log(`\n=== TAKEOVER PROCESSING SUMMARY ===`);
  console.log(`Total drug entries processed: ${totalProcessed}`);
  console.log(`Monthly breakdown:`, monthlyBreakdown);
  console.log(`Takeover Revenue Map entries: ${takeoverRevenueMap.size}`);
  console.log("Takeover Revenue Map sample:", Array.from(takeoverRevenueMap.entries()).slice(0, 10));
  console.log(`Takeover Drug Predictions count: ${takeoverDrugPredictions.length}`);

  // Add takeover drug predictions to the main drug predictions sheet
  if (takeoverDrugPredictions.length > 0) {
    const currentPredictionsRange = drugsTakeoverExpirationPredictions.getUsedRange();
    let currentRowCount = 1;
    if (currentPredictionsRange) {
      currentPredictionsRange.load("rowCount");
      await context.sync();
      currentRowCount = currentPredictionsRange.rowCount;
    }
    
    console.log(`Adding ${takeoverDrugPredictions.length} takeover predictions starting at row ${currentRowCount}`);
    
    drugsTakeoverExpirationPredictions.getRangeByIndexes(
      currentRowCount,
      0,
      takeoverDrugPredictions.length,
      takeoverDrugPredictions[0].length
    ).values = takeoverDrugPredictions;
  } else {
    console.log("No takeover predictions to add!");
  }

} catch (error) {
  console.log("Error processing takeover kit data:", error);
}
// ========== END TAKEOVER KIT CALCULATIONS ==========
      
      //Creating calculation for all drugs per month with monthly price increases
      let newKitDrugPredictions = [];
      Object.keys(salesHistory).forEach((month) => {
        let totalKitAmount = salesHistory[month];
        Object.keys(emkDetails).forEach((kit) => {
          let kitAmount = Math.floor(totalKitAmount * emkDetails[kit].newKitShares);
          if (kitAmount < 1) return;
          emkDetails[kit].drugs.forEach((drug) => {
            if (medsObj[drug].shelfLife == "" || medsObj[drug].shelfLife == "N/A") return;
            
            // Apply monthly increase to drug price
            const drugDate = parseMonthString(month);
            const adjustedDrugPrice = getAdjustedPrice(medsObj[drug].basePrice, drugDate, monthlyIncreaseMap);
            
            newKitDrugPredictions.push([
              month,
              kit,
              drug,
              kitAmount,
              adjustedDrugPrice * kitAmount, // Use adjusted price
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
      
      //Add the Future expiration dates for the auto replenishments with monthly price increases
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

        // Apply monthly increase to the base price
        const adjustedPrice = baseDate ? getAdjustedPrice(price, baseDate, monthlyIncreaseMap) : price;

        // Always include original (keep existing records even if no shelf life)
        outputAutoReplenishAndForecast.push([
          group,
          company,
          medication,
          baseDate ? formatMonth(baseDate) : "N/A",
          adjustedPrice,
          autoReplenish,
          "",
        ]);

        // Only create future replenishments if shelf life exists and is valid
        if (!config || !baseDate || !config.shelfLife || config.shelfLife === "" || config.shelfLife === "N/A" || isNaN(config.shelfLife)){ 
          continue
        };

        const shelfLife = parseInt(config.shelfLife);
        if (shelfLife <= 0) continue;

        for (let i = 1; i <= 20; i++) {
          const futureDate = addDays(baseDate, shelfLife * i);
          const futureAdjustedPrice = getAdjustedPrice(price, futureDate, monthlyIncreaseMap);
          
          outputAutoReplenishAndForecast.push([
            group,
            company,
            medication,
            formatMonth(futureDate),
            parseFloat(futureAdjustedPrice.toFixed(2)),
            autoReplenish,
            "Generated",
          ]);
        }
      }

      // Get AED Auto Replenish groups data
      let wsAutoReplenishAEDGroups = context.workbook.worksheets.getItem("auto_replenish_aed_groups");
      let usedRangeAutoReplenishAEDGroups = wsAutoReplenishAEDGroups.getUsedRange();
      let lastRowAutoReplenishAEDGroups = usedRangeAutoReplenishAEDGroups.getLastRow();
      lastRowAutoReplenishAEDGroups.load("rowIndex");
      await context.sync();
      
      let rangeAutoReplenishAEDGroupsAll = wsAutoReplenishAEDGroups.getRangeByIndexes(
        2,
        0,
        lastRowAutoReplenishAEDGroups.rowIndex - 1,
        6
      );
      rangeAutoReplenishAEDGroupsAll.load("values");
      await context.sync();
      
      console.log("AED Auto Replenish data:", rangeAutoReplenishAEDGroupsAll.values);
      
      // Add AED Auto Replenish data with monthly price increases
      const outputAEDAutoReplenishAndForecast = [
        [
          "Group",
          "Company",
          "AED Item",
          "Expiration",
          "Price",
          "Auto Replenish",
          "Generated Dates",
        ],
      ];

      for (const row of rangeAutoReplenishAEDGroupsAll.values) {
        const [group, company, aedItem, expirationStr, price, autoReplenish] = row;

        if (autoReplenish !== "Enabled") continue;

        // Clean up the AED item name to match the AED_Details sheet
        const cleanAEDItem = aedItem ? aedItem.toString().trim() : '';
        
        const config = aedMedsObj[cleanAEDItem];
        const baseDate = expirationStr && expirationStr !== "N/A" && expirationStr !== "" 
          ? excelSerialDateToJSDate(expirationStr) 
          : null;

        // Apply monthly increase to the base price
        const adjustedPrice = baseDate ? getAdjustedPrice(price, baseDate, monthlyIncreaseMap) : price;

        // Always include original (keep existing records even if no shelf life or expiration)
        outputAEDAutoReplenishAndForecast.push([
          group,
          company,
          cleanAEDItem,
          baseDate ? formatMonth(baseDate) : "N/A",
          adjustedPrice,
          autoReplenish,
          "",
        ]);

        // Only create future replenishments if we have valid expiration date AND shelf life
        if (!baseDate || !config || !config.shelfLife || config.shelfLife === "" || config.shelfLife === "N/A" || isNaN(config.shelfLife)){ 
          continue
        };

        const shelfLife = parseInt(config.shelfLife);
        if (shelfLife <= 0) continue;

        for (let i = 1; i <= 20; i++) {
          const futureDate = addDays(baseDate, shelfLife * i);
          const futureAdjustedPrice = getAdjustedPrice(price, futureDate, monthlyIncreaseMap);
          
          outputAEDAutoReplenishAndForecast.push([
            group,
            company,
            cleanAEDItem,
            formatMonth(futureDate),
            parseFloat(futureAdjustedPrice.toFixed(2)),
            autoReplenish,
            "Generated",
          ]);
        }
      }

      console.log("AED Auto Replenish processed items:", outputAEDAutoReplenishAndForecast.length);

      // Auto-replenish items (only applied once) - Keep separate for different columns
      let autoReplenish = applyAutoReplenishOnce(forecastMap, outputAutoReplenishAndForecast);
      let aedAutoReplenish = applyAutoReplenishOnce(forecastMap, outputAEDAutoReplenishAndForecast);
      
      console.log("Regular Auto Replenish:", autoReplenish);
      console.log("AED Auto Replenish:", aedAutoReplenish);

      // 1. Combine all unique months
      const allMonths = new Set([
        ...drugDataMap.keys(),
        ...autoReplenish.keys(),
        ...aedAutoReplenish.keys(),
        ...baseMap.keys(),
        ...forecastMap.keys(),
        ...takeoverRevenueMap.keys(), // Add takeover months
      ]);

      // 2. Generate final forecast array with separate AED Auto Replenish column
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

      // Sort months chronologically
      const sortedMonths = [...allMonths].sort();
      
      for (const month of sortedMonths) {
        const newkit = baseMap.get(month) || 0;
        const auto = autoReplenish.get(month) || 0;
        const aedAuto = aedAutoReplenish.get(month) || 0;
        const drugData = drugDataMap.get(month) || 0;
        const takeoverRevenue = takeoverRevenueMap.get(month) || 0; // New takeover revenue
        
        // Calculate Non-Auto Replenishment (25% of Auto Replenish, only for future months)
        const nonAutoReplenishment = (month > currentMonthKey) ? auto * nonAutoPercentage : 0;
        
        // Calculate AED Sales (only for future months)
        const aedSales = (month > currentMonthKey) ? aedSalesValue : 0;
        
        // Update total revenue to include all components including takeover revenue
        const totalRevenue = newkit + auto + aedAuto + drugData + nonAutoReplenishment + aedSales + takeoverRevenue;

        finalRevenueForecast.push([
          month, 
          totalRevenue, 
          newkit, 
          auto, 
          drugData,
          nonAutoReplenishment,
          aedSales,
          aedAuto,  // Column H: AED Auto Replenish
          takeoverRevenue  // Column I: New Takeover Kit Revenue
        ]);
      }

      // Update the range to include column I (takeover revenue)
      wsRevenuePredictions.getRangeByIndexes(
        1,
        0,
        finalRevenueForecast.length,
        finalRevenueForecast[0].length
      ).values = finalRevenueForecast;
      
      // Add headers including takeover revenue
      const headers = [
        "Month", 
        "Total Revenue", 
        "New Kit Revenue", 
        "Auto Replenish", 
        "Drug Data Revenue",
        "Non-Auto Replenishment",
        "AED Sales",
        "AED Auto Replenish",
        "New Takeover Kit Revenue"  // New column for takeover revenue
      ];
      wsRevenuePredictions.getRangeByIndexes(0, 0, 1, headers.length).values = [headers];
      
      await context.sync();
      console.log("Final Revenue Forecast:", finalRevenueForecast);
      console.table(finalRevenueForecast);

      // --- Separate History Sheets ---

      // Regular AutoReplenishHistory (only regular auto replenish)
      let autoReplenishDatesAndData = Array.from(autoReplenish.entries()).map(([month, value]) => [month, value]);
      autoReplenishDatesAndData.sort((a, b) => a[0].localeCompare(b[0]));
      
      const sheet = context.workbook.worksheets.getItem("AutoReplenishHistory");
      
      //Get last column
      const autoReplenishHistoryUsedRange = sheet.getUsedRange();
      autoReplenishHistoryUsedRange.load(["columnIndex", "columnCount","rowCount"]);
      await context.sync();
      
      // Load row 2 headers (row index 1)
      if(autoReplenishHistoryUsedRange.columnCount > 1) {
        const headerRange = sheet.getRangeByIndexes(1, 1, 1, autoReplenishHistoryUsedRange.columnCount-1);
        const datesRange = sheet.getRangeByIndexes(2,0,autoReplenishHistoryUsedRange.rowCount-2,1);
        headerRange.load("values");
        datesRange.load("values");
        await context.sync();
        
        console.log("Dates in History:", datesRange.values);
        const headersHistory = headerRange.values[0] // Row 2
        let currentMonthFormatted = formatDateWithDay(new Date());
        let currentMonthPostion = headersHistory.indexOf(currentMonthFormatted);
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
      }

// AED AutoReplenishHistory - CORRECT FORMAT
if (aedAutoReplenish.size > 0) {
  try {
    let aedAutoReplenishDatesAndData = Array.from(aedAutoReplenish.entries()).map(([month, value]) => [month, value]);
    aedAutoReplenishDatesAndData.sort((a, b) => a[0].localeCompare(b[0]));
    
    // Create or get the AED History sheet
    let aedHistorySheet;
    try {
      aedHistorySheet = context.workbook.worksheets.getItem("AutoReplenishAEDHistory");
    } catch (error) {
      // Sheet doesn't exist, create it
      aedHistorySheet = context.workbook.worksheets.add("AutoReplenishAEDHistory");
    }
    
    // Get existing data structure
    const aedHistoryUsedRange = aedHistorySheet.getUsedRange();
    let currentDateFormatted = formatDateWithDay(new Date());
    
    if (!aedHistoryUsedRange) {
      // Initialize new sheet structure
      // Set up headers - "Report Dates" in A1, "Replenish Dates" in A2
      aedHistorySheet.getRange("A1").values = [["Report Dates"]];
      aedHistorySheet.getRange("A2").values = [["Replenish Dates"]];
      
      // Write the months in column A starting from A3
      let allMonths = [...new Set([...aedAutoReplenishDatesAndData.map(x => x[0])])].sort();
      let datesColumn = allMonths.map(month => [month]);
      aedHistorySheet.getRange("A3:A" + (datesColumn.length + 2)).values = datesColumn;
      
      // Write current date as first column header in B2 (NOT B1)
      aedHistorySheet.getRange("B2").numberFormat = [["@"]]; 
      aedHistorySheet.getRange("B2").values = [[currentDateFormatted]];
      console.log("Hello Alvin");
      // Write the AED values in column B starting from B3
      let valuesColumn = allMonths.map(month => {
        return [aedAutoReplenish.get(month) || 0];
      });
      aedHistorySheet.getRange("B3:B" + (valuesColumn.length + 2)).values = valuesColumn;
      
    } else {
      // Existing sheet with data - check if we need to add new column
      aedHistoryUsedRange.load(["columnCount", "rowCount"]);
      await context.sync();
      
      // Get the last column header from ROW 2 to check if it's the same date
      const lastHeaderRange = aedHistorySheet.getRangeByIndexes(1, aedHistoryUsedRange.columnCount - 1, 1, 1);
      lastHeaderRange.load("values");
      await context.sync();
      
      const lastHeaderDate = lastHeaderRange.values[0][0];
      
      // Only add new column if it's a different day
      if (lastHeaderDate !== currentDateFormatted) {
        // Get existing dates column (starting from row 3)
        const datesColumnRange = aedHistorySheet.getRange("A3:A" + aedHistoryUsedRange.rowCount);
        datesColumnRange.load("values");
        await context.sync();
        
        const existingDates = datesColumnRange.values.map(row => row[0]).filter(date => date);
        
        // Add new column for current date
        let newColumnIndex = aedHistoryUsedRange.columnCount;
        
        // Set up new column header in ROW 2
        aedHistorySheet.getRangeByIndexes(1, newColumnIndex, 1, 1).numberFormat = [["@"]];
        aedHistorySheet.getRangeByIndexes(1, newColumnIndex, 1, 1).values = [[currentDateFormatted]];
        
        // Update values for existing dates starting from ROW 3
        for (let i = 0; i < existingDates.length; i++) {
          let month = existingDates[i];
          let value = aedAutoReplenish.get(month) || 0;
          aedHistorySheet.getRangeByIndexes(i + 2, newColumnIndex, 1, 1).values = [[value]];
        }
        
        // Add any new months that don't exist in the dates column
        let newMonths = aedAutoReplenishDatesAndData
          .map(([month]) => month)
          .filter(month => !existingDates.includes(month))
          .sort();
        
        if (newMonths.length > 0) {
          let startRow = existingDates.length + 2;
          
          // Add new months to column A starting from the next available row
          aedHistorySheet.getRangeByIndexes(startRow, 0, newMonths.length, 1).values = newMonths.map(month => [month]);
          
          // Add values for new months in the new column
          for (let i = 0; i < newMonths.length; i++) {
            let month = newMonths[i];
            let value = aedAutoReplenish.get(month) || 0;
            aedHistorySheet.getRangeByIndexes(startRow + i, newColumnIndex, 1, 1).values = [[value]];
          }
        }
        
        console.log("New column added to AED AutoReplenishHistory for date:", currentDateFormatted);
      } else {
        console.log("No new column added - same date as last run:", currentDateFormatted);
      }
    }
    
    await context.sync();
    console.log("AED AutoReplenishHistory updated successfully");
  } catch (error) {
    console.log("Error creating/updating AutoReplenishAEDHistory:", error);
  }
}
  // Add this code right before the final await context.sync() and return context.sync() in your run function
// Place it after you've populated the wsRevenuePredictions sheet

// ========== YEARLY REVENUE SUMMARY ==========
try {
  let yearlyRevenueSheet;
  try {
    yearlyRevenueSheet = context.workbook.worksheets.getItem("Yearly Revenue");
    // Clear existing data
    yearlyRevenueSheet.getUsedRange().clear(Excel.ClearApplyTo.contents);
  } catch (error) {
    // Sheet doesn't exist, create it
    yearlyRevenueSheet = context.workbook.worksheets.add("Yearly Revenue");
  }

  // Get the revenue predictions data (skip header row)
  const revenueDataRange = wsRevenuePredictions.getUsedRange();
  revenueDataRange.load("values");
  await context.sync();
  
  const revenueData = revenueDataRange.values;
  
  // Skip header row (index 0) and process data rows
  const yearlySummary = new Map();
  
  for (let i = 1; i < revenueData.length; i++) {
    const row = revenueData[i];
    if (!row || row.length < 9) continue;
    
    const [month, totalRevenue, newKit, autoReplenish, drugData, nonAuto, aedSales, aedAuto, takeoverRevenue] = row;
    
    // Extract year from month (format: "YYYY-MM")
    const year = month.split('-')[0];
    
    if (!yearlySummary.has(year)) {
      yearlySummary.set(year, {
        total: 0,
        newKit: 0,
        autoReplenish: 0,
        drugData: 0,
        nonAuto: 0,
        aedSales: 0,
        aedAuto: 0,
        takeover: 0
      });
    }
    
    const yearData = yearlySummary.get(year);
    yearData.total += parseFloat(totalRevenue) || 0;
    yearData.newKit += parseFloat(newKit) || 0;
    yearData.autoReplenish += parseFloat(autoReplenish) || 0;
    yearData.drugData += parseFloat(drugData) || 0;
    yearData.nonAuto += parseFloat(nonAuto) || 0;
    yearData.aedSales += parseFloat(aedSales) || 0;
    yearData.aedAuto += parseFloat(aedAuto) || 0;
    yearData.takeover += parseFloat(takeoverRevenue) || 0;
  }
  
  // Convert map to array and sort by year
  const yearlyArray = Array.from(yearlySummary.entries())
    .map(([year, data]) => [year, data.total, data.newKit, data.autoReplenish, data.drugData, data.nonAuto, data.aedSales, data.aedAuto, data.takeover])
    .sort((a, b) => a[0].localeCompare(b[0]));
  
  // Add headers
  const headers = [
    "Year", 
    "Total Revenue", 
    "New Kit Revenue", 
    "Auto Replenish", 
    "Drug Data Revenue",
    "Non-Auto Replenishment",
    "AED Sales",
    "AED Auto Replenish",
    "New Takeover Kit Revenue"
  ];
  
  // Prepare data for Excel (headers + sorted yearly data)
  const sheetData = [headers, ...yearlyArray];
  
  // Write to sheet
  yearlyRevenueSheet.getRangeByIndexes(0, 0, sheetData.length, headers.length).values = sheetData;
  
  // Format the header row
  const headerRange = yearlyRevenueSheet.getRangeByIndexes(0, 0, 1, headers.length);
  headerRange.format.fill.color = "#4472C4";
  headerRange.format.font.color = "white";
  headerRange.format.font.bold = true;
  
  // Format numbers as currency
  if (sheetData.length > 1) {
    const dataRange = yearlyRevenueSheet.getRangeByIndexes(1, 1, sheetData.length - 1, headers.length - 1);
    dataRange.numberFormat = [["$#,##0.00"]];
  }
  
  // Auto-fit columns for better readability
  yearlyRevenueSheet.getUsedRange().format.autofitColumns();
  
  console.log("Yearly Revenue summary created successfully");
  console.table(yearlyArray);
  
} catch (error) {
  console.log("Error creating Yearly Revenue summary:", error);
}
// ========== END YEARLY REVENUE SUMMARY ==========      


      await context.sync();
      return context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

// ... (all helper functions remain exactly the same)

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

// ... (all helper functions remain the same)

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