function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('材料検索')
    .setWidth(1000)
    .setHeight(1400);
}

function searchMaterials(highVoltagePatterns, highVoltageSizes, highVoltageQuantities,
                         lowVoltagePatterns, lowVoltageSizes, lowVoltageQuantities,
                         singlePhasePatterns, singlePhaseSizes, singlePhaseQuantities,
                         lightingPatterns, lightingSizes, lightingQuantities,
                         overheadGroundWirePatterns, overheadGroundWireSizes, overheadGroundWireQuantities,
                         guyWirePatterns, guyWireSizes, guyWireQuantities) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シート2');
  let allResults = [];

  if (highVoltagePatterns.length > 0) {
    allResults = allResults.concat(searchCategoryMaterials(sheet, '高圧', highVoltagePatterns, highVoltageSizes, highVoltageQuantities));
  }
  if (lowVoltagePatterns.length > 0) {
    allResults = allResults.concat(searchCategoryMaterials(sheet, '低圧', lowVoltagePatterns, lowVoltageSizes, lowVoltageQuantities));
  }
  if (singlePhasePatterns.length > 0) {
    allResults = allResults.concat(searchCategoryMaterials(sheet, '単相変台', singlePhasePatterns, singlePhaseSizes, singlePhaseQuantities));
  }
  if (lightingPatterns.length > 0) {
    allResults = allResults.concat(searchCategoryMaterials(sheet, '灯動変台', lightingPatterns, lightingSizes, lightingQuantities));
  }
  if (overheadGroundWirePatterns.length > 0) {
    allResults = allResults.concat(searchCategoryMaterials(sheet, '架空地線', overheadGroundWirePatterns, overheadGroundWireSizes, overheadGroundWireQuantities));
  }
  if (guyWirePatterns.length > 0) {
    allResults = allResults.concat(searchCategoryMaterials(sheet, '支線', guyWirePatterns, guyWireSizes, guyWireQuantities));
  }

  const aggregatedResults = aggregateResults(allResults);
  logSearchUsage();
  return aggregatedResults;
}

function searchCategoryMaterials(sheet, category, patterns, sizes, quantities) {
  const results = [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < patterns.length; i++) {
    const pattern = patterns[i];
    const size = sizes[i];
    const multiplier = quantities[i];

    for (let j = 0; j < data.length; j++) {
      if (data[j][2] === category && data[j][0] === pattern && data[j][1] === size) {
        for (let k = 3; k < data[j].length; k += 2) {
          if (data[j][k] && data[j][k + 1]) {
            results.push({
              material: data[j][k],
              quantity: data[j][k + 1] * multiplier
            });
          }
        }
      }
    }
  }

  return results;
}

function aggregateResults(results) {
  const aggregatedResults = {};
  results.forEach(({ material, quantity }) => {
    if (aggregatedResults[material]) {
      aggregatedResults[material] += quantity;
    } else {
      aggregatedResults[material] = quantity;
    }
  });
  return aggregatedResults;
}

function getDropdownOptionsByCategory(category) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シート2');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();

  const patterns = new Set();
  const sizes = new Set();

  data.forEach(row => {
    if (row[2] === category) {
      patterns.add(row[0]);
      sizes.add(row[1]);
    }
  });

  return {
    patterns: Array.from(patterns),
    sizes: Array.from(sizes)
  };
}

function logSearchUsage() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Log');
    sheet.appendRow(['Timestamp', 'Search Count', 'User']);
  }

  const timestamp = new Date();
  const userEmail = Session.getActiveUser().getEmail();
  const lastRow = sheet.getLastRow();
  const currentCount = lastRow;

  sheet.appendRow([timestamp, currentCount + 1, userEmail]);
}
