import { Component, NgZone } from "@angular/core";
//import * as excel from "./excel.app.component";
//import * as selectionForm from "selection-form/selectionForm.component";
import * as OfficeHelpers from "@microsoft/office-js-helpers";
import { ExcelTableUtil } from "../../../src/utils/excelTableUtils";

const ALPHAVANTAGE_APIKEY = "{{VVI23S5D0LRGBXOJ}}";


@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  /*  welcomeMessage = "Ola";
 
   async run() {
     const excelComponent = new excel.default();
     return excelComponent.run();
   }
   async runSelectionForm() {
     const selectionForm = new selectionForm.de();
     return selectionForm.run();
   } */
  symbols = [];
  error = null;
  waiting = false;
  zone = new NgZone({});

  tableUtil = new ExcelTableUtil("Portfolio", "A1:H1", [
    "Symbol",
    "Last Price",
    "Timestamp",
    "Quantity",
    "Price Paid",
    "Total Gain",
    "Total Gain %",
    "Value",
  ]);

  constructor() {
    this.syncTable().then(() => { });
  }

  // Adds symbol
  addSymbol = async (symbol) => {
    this.waiting = true;

    // Get quote and add to Excel table
    this.getQuote(symbol).then(
      (res) => {
        let cnt = this.symbols.length;
        const data = [
          res["01. symbol"], //Symbol
          res["05. price"], //Last Price
          res["07. latest trading day"], // Timestamp of quote,
          0, // quantity (manually entered)
          0, // price paid (manually entered)
          `=(B${cnt + 2} * D${cnt + 2}) - (E${cnt + 2} * D${cnt + 2})`, //Total Gain $
          `=H${cnt + 2} / (E${cnt + 2} * D${cnt + 2}) * 100 - 100`, //Total Gain %
          `=B${cnt + 2} * D${cnt + 2}`, //Value
        ];
        this.tableUtil.addRow(data).then(
          () => {
            this.symbols.unshift(symbol.toUpperCase());
            this.waiting = false;
          },
          (err) => {
            this.error = err;
          }
        );
      },
      (err) => {
        this.error = err;
        this.waiting = false;
      }
    );
  };

  // Delete symbol
  deleteSymbol = async (index) => {
    // Delete from Excel table by index number
    const symbol: string = this.symbols[index];
    this.waiting = true;
    this.tableUtil.getColumnData("Symbol").then(
      async (columnData: string[]) => {
        // Ensure the symbol was found in the Excel table
        if (columnData.indexOf(symbol) !== -1) {
          this.tableUtil.deleteRow(columnData.indexOf(symbol)).then(
            async () => {
              this.symbols.splice(index, 1);
              this.waiting = false;
            },
            (err) => {
              this.error = err;
              this.waiting = false;
            }
          );
        } else {
          this.symbols.splice(index, 1);
          this.waiting = false;
        }
      },
      (err) => {
        this.error = err;
        this.waiting = false;
      }
    );
  };

  // Refresh symbol
  refreshSymbol = async (index) => {
    // Refresh stock quote and update Excel table
    const symbol = this.symbols[index];
    this.waiting = true;
    this.tableUtil.getColumnData("Symbol").then(
      async (columnData: string[]) => {
        // Ensure the symbol was found in the Excel table
        const rowIndex = columnData.indexOf(symbol);
        if (rowIndex !== -1) {
          this.getQuote(symbol).then((res) => {
            // "last trade" is in column B with a row index offset of 2 (row 0 + the header row)
            this.tableUtil.updateCell(`B${rowIndex + 2}:B${rowIndex + 2}`, res["05. price"]).then(
              async () => {
                this.waiting = false;
              },
              (err) => {
                this.error = err;
                this.waiting = false;
              }
            );
          });
        } else {
          this.error = `${symbol} not found in Excel`;
          this.symbols.splice(index, 1);
          this.waiting = false;
        }
      },
      (err) => {
        this.error = err;
        this.waiting = false;
      }
    );
  };

  // Reads symbols from an existing Excel workbook and pre-populates them in the add-in
  syncTable = async () => {
    this.waiting = true;
    this.tableUtil.getColumnData("Symbol").then(
      async (columnData: string[]) => {
        this.symbols = columnData;
        this.waiting = false;
      },
      (err) => {
        this.error = err;
        this.waiting = false;
      }
    );
  };

  // Gets a quote by calling into the stock service
  getQuote = async (symbol) => {
    return new Promise((resolve, reject) => {
      const queryEndpoint = `https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=${escape(
        symbol
      )}&apikey=${ALPHAVANTAGE_APIKEY}`;

      fetch(queryEndpoint)
        .then((res) => {
          if (!res.ok) {
            reject("Error getting quote");
          }
          return res.json();
        })
        .then((jsonResponse) => {
          const quote = jsonResponse["Global Quote"];
          resolve(quote);
        });
    });
  };
}
