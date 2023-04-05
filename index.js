/**
 * Converts an excel spreadsheet to the json required for api spec fields
 * To run: npm start
 */

const excelToJson = require("convert-excel-to-json");
const fs = require("fs");

/**
 * Converts a string to camel case
 * @param str the string to convert to camel case
 * @returns
 */
function toCamelCase(str) {
  if (!str) {
    return "";
  }
  return str
    .replace(/(?:^\w|[A-Z]|\b\w)/g, function (word, index) {
      return index === 0 ? word.toLowerCase() : word.toUpperCase();
    })
    .replace(/\s+/g, "");
}

const result = excelToJson({
  sourceFile: "C:/Users/corwi/Downloads/injuryClaim.xlsx", // Update file path as required
  header: {
    rows: 3,
  },
  columnToKey: {
    A: "desc",
    B: "sample",
  },
});

/**
 * Insert excel tab names here
 */
const tabNames = ["Injury Claim", "Supplier Claim"];

/**
 * Loop the tabs, convert the data to an object and write to a json file
 */
tabNames.forEach((tabName) => {
  const rows = result[tabName].filter((row) => !!row.sample);

  const obj = {
    fields: [],
  };

  for (let i = 0; i < rows.length - 1; i++) {
    obj.fields.push({ name: toCamelCase(rows[i].desc), value: rows[i].sample });
  }

  const json = JSON.stringify(obj);

  fs.writeFile(`${tabName}_fields.json`, json, "utf8", () => {
    console.log("done");
  });
});
