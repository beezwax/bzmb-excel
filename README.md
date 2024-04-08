# Introduction

A [bzBond-server](https://github.com/beezwax/bzBond/tree/main/packages/bzBond-server#bzbond-server) Microbond to create formatted Excel (.xlsx) documents.

# Installation

## Installation on macOS/Linux

On macOs/Linux use the following command to install this Microbond:

`/var/www/bzbond-server/bin/install-microbond.sh bzmb-excel https://github.com/beezwax/bzmb-excel`

## Installation on Windows Server

On Windows Server use the following command to install this Microbond:

`powershell -File "C:\Program Files\bzBond-server\bin\install-microbond.ps1" bzmb-excel https://github.com/beezwax/bzmb-excel`

## Installation with a proxy on macOS/Linux

On macOs/Linux use the following command to install this Microbond via a proxy:

`/var/www/bzbond-server/bin/install-microbond.sh bzmb-excel https://github.com/beezwax/bzmb-excel http://proxy.example.com:443`

## Installation with a proxy on Windows Server

On Windows Server use the following command to install this Microbond via a proxy:

`powershell -File "C:\Program Files\bzBond-server\bin\install-microbond.ps1" -Proxy http://proxy.example.com:443`

# Usage

The bzmb-excel Microbond provides one route

## bzmb-excel-createWorkbook

In a server-side FileMaker script run `bzBondRelay` script with parameters in the following format:

```
{
  "mode": "PERFORM_JAVASCRIPT",

  "route": "bzmb-excel-createWorkbook",

  "customMethod": "POST",

  "customBody": {
    
    // Required. Array of sheets to include in the workbook 
    "sheets": [
      {
        // Required. Sheet name
        "name": "string",

        // Array of column headers
        "headers": array,
        
        // example headers
        "headers": ["Column 1", "Column 2", "Column 3"]

        // An array or arrays or an array of objects representing rows in a spreadsheet
        // Can also be a combination. Can include style and format references
        "data": array

        // Example array of arrays
        "data": [
          ["A", "B", 3],
          [4, C, D]
        ]

        // Example array of objects
        "data": [
          {
            "Column 3": 3,
            "Column 2": "B",
            "Column 1": "A"
          },
          {
            "Column 1": 4,
            "Column 2": "C",
            "Column 3": "D"
          }
        ]

        // Example array of objects and arrays
        "data": [
          {
            "Column 1": "A",
            "Column 2": "B",
            "Column 3": 3
          },
          [4, C, D]
        ]

        // Example array of objects and arrays with styling
        "data": [
          {
            "Column 1": "A",
            "Column 2": "B",
            "Column 3": {value: "3", style: "percent0dp"}
          },
          [{"value": 4}, {"value: "C", "style": "alignRight"}, {"Column 3": {"value": "D"}}]
        ]
      }
    ],

    // An object of styles that can be referenced by cells
    "styles": {
      "percent0dp": {
        "number": "0%",
      },
      "alignRight": {
        "alignment": {
          "horizontal": "right"
        }
      }
    }
  }
}

```

A base64 represenation of the excel file can be accessed via `Get ( ScriptResult )`:
`JSONGetElement ( Get ( ScriptResult ); "response.result" )`