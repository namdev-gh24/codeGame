/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
require("xlsx");
let file;
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
      fileReader.readAsText(file);
      fileReader.onload = (event) => {
        result = event.target.result;
        run();
        showFields();
      };
    });

    document.getElementById("run").onclick = run;
    document.getElementById("show").onclick = showFields;
    document.getElementById("ai").onclick = genAI;
  }
});

export function genAI() {
  console.log("Calling AI To Generate Random Names");
  // fetch("https://randomuser.me/api/?results=4")
  //   .then((results) => {
  //     return results.json();
  //   })
  //   .then((data) => {
  //     console.log(data.results);
  //     // Access your data here
  //   });
  const apiKey = "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";

  // Set the API endpoint URL
  const apiUrl = "https://api.openai.com/v1/engines/davinci/completions";

  // Set the request headers
  const headers = {
    "Content-Type": "application/json",
    Authorization: `Bearer ${apiKey}`,
  };

  // Set the request body
  const body = {
    prompt: "List of vehicle names:\n- Car\n- Truck\n- Bus\n- Motorcycle\n- Bicycle\n",
    temperature: 0.5,
    max_tokens: 5,
    n: 5,
    stop: "\n",
  };

  // Send the API request using jQuery
  $.ajax({
    type: "POST",
    url: apiUrl,
    headers: headers,
    data: JSON.stringify(body),
    success: function (response) {
      // Parse the response JSON
      const choices = response.choices; // Extract the text from each choice
      const vehicleNames = choices.map((choice) => choice.text.trim()); // Log the array of vehicle names
      console.log(vehicleNames);
    },
    error: function (xhr, status, error) {
      // Log an error message if the request failed
      console.error("Request failed. Returned status of " + xhr.status + ". Error message: " + error);
    },
  });
}

export function genRandString(length) {
  let result = "";
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

export function genRandDate(start, end) {
  return new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
}

export function dropdownEventHandler(dropDownEvent) {
  var element = dropDownEvent.currentTarget.id.slice(-1);
  var checkbox_element = document.getElementById("checkboxw" + element);
  checkbox_element.checked = false;
  checkbox_element.dispatchEvent(new Event("change"));
  dropdownValue = dropDownEvent.target.value;
}

export function drop2downEventHandler(drop2DownEvent) {
  var element = drop2DownEvent.currentTarget.id.slice(-1);
  var checkbox_element = document.getElementById("checkboxw" + element);
  checkbox_element.checked = false;
  checkbox_element.dispatchEvent(new Event("change"));
  manipulatedValue = drop2DownEvent.target.value;
}

export function checkBoxEventHandler(checkBoxEvent) {
  var element = checkBoxEvent.currentTarget.id.slice(-1);
  if (checkBoxEvent.target.checked) {
    selected.splice(selected.length, 0, { [checkBoxEvent.target.name]: [dropdownValue, manipulatedValue, element] });
    dropdownValue = "string";
    manipulatedValue = "genRand";
  } else {
    let y = selected.findIndex(() => {
      for (let i = 0; i < selected.length; i++) {
        return Object.values(selected[i])[0][2] == element;
      }
    });
    selected.splice(y, 1);
  }
}

// created dynamic checkboxes for fields
export function showFields() {
  json = JSON.parse(result);
  let data = Object.keys(json[0]);
  selected = [];
  run();
  document.getElementById("cb").innerHTML = "";

  const dropdownData = {
    string: "String",
    num: "Numbers",
    date: "Date",
  };

  const drop2DownData = {
    genRand: "Generate Random",
    genAI: "Generate By AI",
    encFile: "Encrypt Data",
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
    document.getElementById("dropdown" + i).addEventListener("change", dropdownEventHandler);
    document.getElementById("drop2down" + i).addEventListener("change", drop2downEventHandler);
    document.getElementById("checkboxw" + i).addEventListener("change", checkBoxEventHandler);
  }
}

export function getType(selected) {
  var result = [];
  for (var i = 0; i < selected.length; i++) {
    var selectedValues = Object.values(selected)[i];
    var selectObject = Object.keys(selectedValues)[0];
    var type = Object.values(selectedValues)[0][0];
    var manipulate = Object.values(selectedValues)[0][1];
    result.push([selectObject, type, manipulate]);
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
      for (var i = 0; i < json.length; i++) {
        for (var j = 0; j < selectData.length; j++) {
          if (selectData[j][1] == "string" && selectData[j][2] == "genRand") {
            json[i][selectData[j][0]] = genRandString(7);
          }
          if (selectData[j][1] == "num" && selectData[j][2] == "genRand") {
            json[i][selectData[j][0]] = genRandNum(5);
          }
          if (selectData[j][1] == "date" && selectData[j][2] == "genRand") {
            json[i][selectData[j][0]] = genRandDate(new Date(2012, 0, 1), new Date());
          }
          if (selectData[j][2] == "encFile") {
            json[i][selectData[j][0]] = md5(json[i][selectData[j][0]]);
          }
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
