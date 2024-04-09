const xl = require('excel4node');

const createWorkbook = async (payload) => {

  const {sheets, styles} = payload

  const workbook = new xl.Workbook();
  const styleMap =
    styles ?
      Object.keys(styles).reduce((acc, curr) => 
        ({...acc, ...{[curr]: workbook.createStyle(styles[curr])}}),
        {}
      ) : {};

  sheets.forEach(({headers, data, name, options, columnWidths}, index) => {
    let columnContentLengths = {};
    let rowNumber = 1;
    let headerValues;
    const worksheet = workbook.addWorksheet(name, options);
    if (headers) {
      headerValues = [];
      headers.forEach((header, index) => {
        // Handle object
        let value;
        let style;
        let blank;
        if (typeof header === "object") {
          value = header.value;
          style = header.style;
          blank = header.blank;
        } else {
          value = header;
        }
        const type = typeof value;
        if (!blank) {
          column = index + 1;
          columnContentLengths[column] = columnContentLengths[column] ? [...columnContentLengths[column], ...[value.toString().length]] : [value.toString().length];
          if (styleMap[style] !== undefined) {
            worksheet.cell(rowNumber, column)[type](value).style(styleMap[style]);
          } else {
            worksheet.cell(rowNumber, column)[type](value);
          }
        }
        headerValues = [...headerValues, ...[value]];
      });
      rowNumber++;
    }

    data.forEach(row => {
      let rowType;
      if (Array.isArray(row)) {
        rowType = "array";
      } else if (typeof row === "object"){
        rowType = "object";
      } else {
        throw new Error("Row must be an array or an object");
      }

      if (rowType === "array") {
        row.forEach((cell, index) => {
          let value;
          let style;
          let headerReference;
          if (typeof cell === "object") {
            if ("value" in cell) {
              value = cell.value;
              style = cell.style;
            } else {
              headerReference = Object.keys(cell)[0];
              value = cell[headerReference].value;
              style = cell[headerReference].style;
            }
          } else {
            value = cell;
          }
          const type = typeof value;
          let column = headerValues?.findIndex(headerValue => headerValue === headerReference) + 1;
          if (!column) {
            column = index + 1; 
          }
          if (styleMap[style] !== undefined) {
            worksheet.cell(rowNumber, column)[type](value).style(styleMap[style]);
          } else {
            worksheet.cell(rowNumber, column)[type](value);
          }
          columnContentLengths[column] = columnContentLengths[column] ? [...columnContentLengths[column], ...[value.toString().length]] : [value.toString().length];
        });
      } else if (rowType === "object") {
        Object.keys(row).forEach((cell, index) => {
          let value;
          let style;
          let headerReference = cell;
          if (typeof row[cell] === "object") {
            value = row[cell].value;
            style = row[cell].style;
          } else {
            value = row[cell];
          }
          const type = typeof value;
          let column = headerValues?.findIndex(headerValue => headerValue === headerReference) + 1;
          if (!column) {
            column = index + 1; 
          }
          if (style) {
            worksheet.cell(rowNumber, column)[type](value).style(styleMap[style]);
          } else {
            worksheet.cell(rowNumber, column)[type](value);
          }
          columnContentLengths[column] = columnContentLengths[column] ? [...columnContentLengths[column], ...[value.toString().length]] : [value.toString().length];
        });
      }
      rowNumber++;
    });

    if(columnWidths?.length) {
      columnWidths.forEach(({column, width, adjust = 0}) => {
        let columnWidth;
        if (typeof width === "number") {
          columnWidth = width;
        } else if (width === "autofit") {
          columnWidth = Math.max(...columnContentLengths[column]);
        }
        worksheet.column(column).setWidth(columnWidth + adjust);
      })
    }
  });

  const base64 = (await workbook.writeToBuffer()).toString("base64").replaceAll("\n", "");

  return base64;

}

module.exports = createWorkbook;