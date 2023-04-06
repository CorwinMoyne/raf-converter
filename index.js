/**
 * Converts an excel spreadsheet to the json required for api spec fields
 * To run: npm start
 */

const excelToJson = require("convert-excel-to-json");
const fs = require("fs");

function formatString(str) {
  this.str = str;

  this.cleanUp = function () {
    this.str = this.str
      .replace(/ *\([^)]*\) */g, "") // Remove parenthesis
      .replace(/\./g, "") // Remove period
      .replace(/[^\x00-\x7F]/g, "") // Remove non ascii
      .replace(/\?/g, "") // Remove question mark
      .replace(/\*/g, "") // Remove asterik
      .replace(/\'/g, "") // Remove apostrophie
      .replace(/\//g, ""); // Remove forward slash
    return this;
  };

  this.removeNumbers = function () {
    this.str = this.str.replace(/[0-9]/g, "");
    return this;
  };

  this.firstLetterToLowerCase = function () {
    this.str = this.str.charAt(0).toLowerCase() + this.str.slice(1);
    return this;
  };

  this.trim = function () {
    this.str = this.str.trim();
    return this;
  };

  this.toCamelCase = function () {
    this.str = this.str
      .replace(/(?:^\w|[A-Z]|\b\w)/g, function (word, index) {
        return index === 0 ? word.toLowerCase() : word.toUpperCase();
      })
      .replace(/\s+/g, "");
    return this.str;
  };
}

const result = excelToJson({
  sourceFile: "C:/Users/corwi/Downloads/injuryClaim.xlsx", // Update file path as required
  header: {
    rows: 1,
  },
  columnToKey: {
    A: "desc",
    B: "sample",
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
  const rows = result[tabName];

  let prepend = "";
  let headerIndex = 1;

  const obj = {
    fields: [],
  };

  rows.forEach((row) => {
    if (!row.sample && row.desc.includes(`${headerIndex}.`)) {
      prepend = new formatString(row.desc)
        .cleanUp()
        .removeNumbers()
        .trim()
        .firstLetterToLowerCase()
        .toCamelCase();
      headerIndex++;
    }
    if (row.sample) {
      obj.fields.push({
        name: new formatString(`${prepend}${row.desc}`).cleanUp().toCamelCase(),
        value: row.sample,
      });
    }
  });

  const json = JSON.stringify(obj);

  fs.writeFile(`${tabName}_fields.json`, json, "utf8", () => {
    console.log("done");
  });
});
