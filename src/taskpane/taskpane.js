/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
require("xlsx");
let file;
let vehicleChecked;
let json;
let result;
let dropdownValue = "string";
let manipulatedValue = "genRand";
let selected = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("myFile").addEventListener("change", (Event) => {
      file = Event.target.files[0];
      let fileReader = new FileReader();
      fileReader.readAsBinaryString(file);
      fileReader.onload = (event) => {
        result = event.target.result;
        let workbook = XLSX.read(result, { type: "binary" });
        workbook.SheetNames.forEach((sheet) => {
          let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
        });
      };
    });

    document.getElementById("run").onclick = run;
    document.getElementById("show").onclick = showFields;
    // document.getElementById("ai").onclick = callAI2;
  }
});

// export async function importData() {
//   json = JSON.parse(result);
//   try {
//     await Excel.run(async (context) => {
//       let ws = context.workbook.worksheets.getActiveWorksheet();
//       let range = ws.getRange();
//       console.log("RANGE IS " + range.cellCount.length);
//       await context.sync();
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

export function genAI() {
  console.log("Calling AI To Generate Random Names");
  fetch("https://randomuser.me/api/?results=4")
    .then((results) => {
      return results.json();
    })
    .then((data) => {
      console.log(data.results);
      // Access your data here
    });
}

export function genRandString(length) {
  let result = " ";
  const characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
  const charactersLength = characters.length;
  for (let i = 0; i < length; i++) {
    result += characters.charAt(Math.floor(Math.random() * charactersLength));
  }
  return result;
}

export function genRandNum(length) {
  let result = 0;
  let max = Math.pow(10, length + 1);
  result = Math.floor(100000 + Math.random() * max);
  return result;
}

export function createSelected(dropDownEvent, manipulateEvent, checkBoxEvent) {
  console.log(dropDownEvent.target.value);
  console.log(manipulateEvent.target.value);
  console.log(checkBoxEvent.target.name);
}
// created dynamic checkboxes for fields
export function showFields() {
  json = JSON.parse(result);
  let data = Object.keys(json[0]);

  const dropdownData = {
    string: "String",
    num: "Numbers",
    date: "Date",
  };
  const drop2DownData = {
    genRand: "Generate Random",
    genAI: "Generate By AI",
    genFile: "Generate From File",
  };

  for (var i = 0; i < data.length; i++) {
    var container = document.createElement("div");
    var checkbox = document.createElement("input");

    checkbox.type = "checkbox";
    checkbox.name = data[i];
    checkbox.value = data[i];
    checkbox.id = "checkboxw" + i;

    var dropdown = document.createElement("select");
    dropdown.id = "dropdown" + i;
    for (let data in dropdownData) {
      let option = document.createElement("option");
      option.setAttribute("value", data);
      let optiontext = document.createTextNode(dropdownData[data]);
      option.appendChild(optiontext);
      dropdown.appendChild(option);
    }

    var drop2down = document.createElement("select");
    drop2down.id = "drop2down" + i;
    for (let data in drop2DownData) {
      let option = document.createElement("option");
      option.setAttribute("value", data);

      let optiontext = document.createTextNode(drop2DownData[data]);
      option.appendChild(optiontext);
      drop2down.appendChild(option);
    }

    var label = document.createElement("label");
    label.htmlFor = "checkboxw";
    label.appendChild(document.createTextNode(data[i]));

    container.appendChild(checkbox);
    container.appendChild(label);
    container.appendChild(dropdown);
    container.appendChild(drop2down);
    document.getElementById("cb").appendChild(container);
  }
  for (var i = 0; i < data.length; i++) {
    document.getElementById("dropdown" + i).addEventListener("change", (Event) => {
      dropdownValue = Event.target.value;
      console.log(dropdownValue);
    });
    document.getElementById("drop2down" + i).addEventListener("change", (Event) => {
      manipulatedValue = Event.target.value;
    });

    document.getElementById("checkboxw" + i).addEventListener("change", (Event) => {
      if (Event.target.checked) {
        selected.push({ [Event.target.name]: [dropdownValue, manipulatedValue] });
        dropdownValue = "string";
        manipulatedValue = "genRand";
        // selected.push(Event.target.name);
      } else {
        selected.pop({ [Event.target.name]: [dropdownValue, manipulatedValue] });
        // selected.pop(Event.target.name);
      }
    });
  }
}
export function getType(selected) {
  console.log(selected);
  var result = [];
  for (var i = 0; i < selected.length; i++) {
    var selectedValues = Object.values(selected)[i];
    var selectObject = Object.keys(selectedValues)[0];
    var type = Object.values(selectedValues)[0][0];
    var manipulate = Object.values(selectedValues)[0][1];
    result.push([selectObject, type, manipulate]);
    console.log("RESULT IS - " + result);
  }
  return result;
}

export async function run() {
  var selectData = getType(selected);
  json = JSON.parse(result);
  try {
    await Excel.run(async (context) => {
      // setup workbook and sheet column heading
      let ws = context.workbook.worksheets.getActiveWorksheet();
      let range = ws.getRange("A1:D1");
      let range2 = ws.getRange("A2:D8");
      let data = new Array(Object.keys(json[0]));
      range.values = data;

      range.format.autofitColumns();

      //manipulating json data
      console.log("before for loop manupilatin" + selectData);
      for (var i = 0; i < json.length; i++) {
        for (var j = 0; j < selectData.length; j++) {
          if (selectData[j][1] == "string" && selectData[j][2] == "genRand") {
            json[i][selectData[j][0]] = genRandString(7);
          }
          if (selectData[j][1] == "num" && selectData[j][2] == "genRand") {
            json[i][selectData[j][0]] = genRandNum(5);
          }
          // json[i][selectData[j][0]] = "lorem ipsum";
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
      // console.log(ws);
    });
  } catch (error) {
    console.error(error);
  }
}
