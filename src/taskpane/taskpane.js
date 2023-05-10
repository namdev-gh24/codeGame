/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
require('xlsx');
let file;
let vehicleChecked;
let json;
let result;
let selected = [];


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    document.getElementById("myFile").addEventListener("change", (Event) => {
      file = Event.target.files[0];
      let fileReader = new FileReader();
      fileReader.readAsBinaryString(file);
      //fileReader.readAsArrayBuffer(file);
      fileReader.onload = (event) => {
        console.log(event);
        result = event.target.result;
        let workbook = XLSX.read(result, { type: "binary" });
        console.log(workbook);
        workbook.SheetNames.forEach(sheet => {
          let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
          console.log(rowObject);
        });
      }

    });

    document.getElementById("run").onclick = run;
    document.getElementById("show").onclick = showFields;
  }
});

// created dynamic checkboxes for fields
export function showFields() {
  json = JSON.parse(result);
  let data = Object.keys(json[0]);

  for (var i = 0; i < data.length; i++) {
    var container = document.createElement('div');
    var checkbox = document.createElement('input');
    checkbox.type = "checkbox";
    checkbox.name = data[i];
    checkbox.value = data[i];
    checkbox.id = "checkboxw" + i;

    var label = document.createElement('label')
    label.htmlFor = "checkboxw";
    label.appendChild(document.createTextNode(data[i]));

    container.appendChild(checkbox);
    container.appendChild(label);
    document.getElementById("cb").appendChild(container);
  }
  for (var i = 0; i < data.length; i++) {
    document.getElementById("checkboxw" + i).addEventListener("change", (Event) => {
      console.log(Event);
      if (Event.target.checked) {
        selected.push(Event.target.name);
      }
      else {
        console.log(Event.target.name + "  POPPED");
        selected.pop(Event.target.name);
      }

    });
  }


}

export async function run() {
  json = JSON.parse(result);
  try {

    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      // setup workbook and sheet column heading
      let ws = context.workbook.worksheets.getActiveWorksheet();
      let range = ws.getRange("A1:D1");
      let range2 = ws.getRange("A2:D8");
      let data = new Array(Object.keys(json[0]));
      range.values = data;
      range.format.autofitColumns();

      //manipulating json data
      console.log("before for loop manupilatin" + selected[0]);
      for (var i = 0; i < json.length; i++) {
        for (var j = 0; j < selected.length; j++) {
          json[i][selected[j]] = "lorem ipsum";

        }
      }




      // add data to excel from json
      let data3 = [];
      for (let i = 0; i < json.length; i++) {
        let data2 = [];
        data2.push(json[i].Vehicle);
        data2.push(json[i].Date);
        data2.push(json[i].Location);
        data2.push(json[i].Speed);

        data3.push(data2);

      }
      range2.values = data3;
      range2.format.autofitColumns();

      await context.sync();
      console.log(ws);


    });
    parseData();
  } catch (error) {
    console.error(error);
  }

  function parseData() {
    // Get the current worksheet
    var sheet = Office.context.document.workbook.worksheets.getActiveWorksheet();

    // Define the dynamic range of cells to parse
    var range = sheet.getUsedRange();

    // Load the values of the cells in the range
    range.load("values");

    // Run a batch operation to get the cell values
    Office.context.document.batch(function (batch) {
      batch.sync();
      // Access the cell values after the batch operation has completed
      var cellValues = range.values;
      // Parse the data as needed
      // ...
      console.log(cellValues);
    });
  }
}
