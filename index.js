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
    C: "dataFormat",
    E: "required",
  },
});

/**
 * Insert excel tab names here
 */
const tabNames = ["Injury Claim", "Death Claim"];

/**
 * Loop the tabs, convert the data to an object and write to a json file
 */
tabNames.forEach((tabName) => {
  const rows = result[tabName].filter((row) => !!row.sample);

  const obj = {
    fields: [],
  };

  rows.forEach((row) =>
    obj.fields.push({
      name: toCamelCase(row.desc),
      value: row.sample,
      in: "path",
      description: row.desc,
      required: row.required,
      schema: {
        type: row.desc.includes("no.") ? "number" : "string",
        format: row.dataFormat,
      },
    })
  );

  const json = JSON.stringify(obj);

  fs.writeFile(`${tabName}_fields.json`, json, "utf8", () => {
    console.log("done");
  });
});
