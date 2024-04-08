const createWorkbook = require("./createWorkbook.js");
const fs = require("fs");
const crypto = require("crypto");

const args = process.argv.slice(2);

const params = [
  {
    sheets: [{
      "name": "Sheet1",
      "headers": ["Column1", "Column2", "Column3"],
      "data" : [
        ["A", "B", "C"],
        {"Column1": "A", "Column2": {"value": "B", "style": "rightAlign"}, "Column3": "C"},
        {"Column2": "A", "Column1": "B", "Column3": "C"},
        ["A", {"value": 1, "style": "percent0dp"}, "C"]
      ],
      "options": "{}"
    }],
  
    styles: {
      "percent0dp": {
        "numberFormat": "0%"
      },
      "rightAlign": {
        "alignment": {
          "horizontal": "right"
        }
      }
    },

    filename: "header_and_formatting.xlsx"
  },
  {
    sheets: [{
      "name": "Sheet1",
      "data" : [
        ["A", "B", "C"],
        ["A", {"value": "B", "style": "rightAlign"}, "C"],
        {"Column2": "A", "Column1": "B", "Column3": "C"},
        ["A", {"value": 1, "style": "percent0dp"}, "C"]
      ]
    }],

    filename: "no_header_missing_styles.xlsx"
  },
  {
    sheets: [{
      "name": "Sheet1",
      "data" : [
        ["A", "B", "C"],
        ["D", "E", "F"],
        ["1", 1, 2],
        [{value: 1}, "A", 1]
      ]
    }],

    filename: "no_header_no_formatting.xlsx"
  }
];

(async () => {
  params.forEach(async ({sheets, styles, filename}) => {
    const base64 = await createWorkbook({sheets, styles});
    fs.writeFileSync(`test_output/${filename}`, base64, {encoding: "base64"});
  });
})();