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
let manipulatedData = [];
let selected = [];
let rows = [];
let columns;
let dropdownData = {
  string: "String",
  num: "Numbers",
  date: "Date",
};

let drop2DownData = {
  genRand: "Generate Random",
  genAI: "Generate By AI",
  encFile: "Encrypt Data",
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("myFile").addEventListener("change", (Event) => {
      file = Event.target.files[0];
      var ext = file.name.split(".")[1];
      let fileReader = new FileReader();
      fileReader.readAsText(file);
      fileReader.onload = (event) => {
        result = event.target.result;
        if (ext == "json") {
          loadJSONFile();
        } else if (ext == "xml") {
          loadXMLFile();
        } else {
          console.error("[-] Choose Appropriate File");
        }
      };
    });
    document.getElementById("run").onclick = manipulateData;
    document.getElementById("show").onclick = showFields;
    document.getElementById("ai").onclick = genAI;
    document.getElementById("exportToXML").addEventListener("click", () => {
      function OBJtoXML(obj) {
        var xml = "";
        for (var prop in obj) {
          xml += "<" + prop + ">";
          if (obj[prop] instanceof Array) {
            for (var array in obj[prop]) {
              xml += OBJtoXML(new Object(obj[prop][array]));
            }
          } else if (typeof obj[prop] == "object") {
            xml += OBJtoXML(new Object(obj[prop]));
          } else {
            xml += obj[prop];
          }
          xml += "</" + prop + ">";
        }
        var xml = xml.replace(/(<\/[0-9]>)+/g, "</" + "tag" + ">");
        xml = xml.replace(/(<[0-9]>)+/g, "<" + "tag" + ">");
        return xml;
      }
      var xmltext = "<body>" + OBJtoXML(json) + "</body>";
      var a = document.createElement("a");
      a.href = window.URL.createObjectURL(new Blob([xmltext], { type: "text/xml" }));
      a.download = "demo.xml";
      a.click();
    });

    document.getElementById("exportToJSON").addEventListener("click", () => {
      var a = document.createElement("a");
      var file = new Blob([JSON.stringify(rows)], { type: "text/plain" });
      a.href = URL.createObjectURL(file);
      a.download = "demo.json";
      a.click();
    });
  }
});

export function OBJtoXML() {}

export async function loadJSONFile() {
  json = JSON.parse(result);
  try {
    await Excel.run(async (context) => {
      let ws = context.workbook.worksheets.getActiveWorksheet();
      let range = ws.getRange("A1:D1");
      let range2 = ws.getRange("A2:D8");
      let data = new Array(Object.keys(json[0]));
      range.values = data;
      range.format.autofitColumns();

      let data3 = [];
      for (let i = 0; i < json.length; i++) {
        data3.push(Object.values(json[i]));
      }
      range2.values = data3;
      range2.format.autofitColumns();

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
  showFields();
}

export function loadXMLFile() {
  console.error("EXCEL SE IMPORT KARLE BHAI");
}

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
  const apiKey = "SECRET_KEY";

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
  console.log(element);
  if (checkBoxEvent.target.checked) {
    document.getElementById("dropdown" + element).disabled = false;
    document.getElementById("dropdown" + element).hidden = false;
    document.getElementById("drop2down" + element).hidden = false;
    document.getElementById("drop2down" + element).disabled = false;
    selected.splice(selected.length, 0, { [checkBoxEvent.target.name]: [dropdownValue, manipulatedValue, element] });
    dropdownValue = "string";
    manipulatedValue = "genRand";
    console.log("CHECKED - " + selected);
  } else {
    document.getElementById("dropdown" + element).disabled = true;
    document.getElementById("dropdown" + element).hidden = true;
    document.getElementById("drop2down" + element).hidden = true;
    document.getElementById("drop2down" + element).disabled = true;
    let y = selected.findIndex(() => {
      for (let i = 0; i < selected.length; i++) {
        return Object.values(selected[i])[0][2] == element;
      }
    });
    console.log(y);
    console.log("UNCHECKED BEFORE - " + Object.values(selected));
    if (y != -1) {
      selected.splice(y, 1);
    }
    console.log("UNCHECKED AFTER - " + Object.values(selected));
  }
}

// created dynamic checkboxes for fields
export async function showFields() {
  selected = [];
  document.getElementById("cb").innerHTML = "";

  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    var selectedRange = sheet.getUsedRange();
    selectedRange.load("values");

    await context.sync().then(function () {
      json = selectedRange.values;
      columns = json[0];
      for (var i = 1; i < json.length; i++) {
        var row = {};
        for (var j = 0; j < json[i].length; j++) {
          row[columns[j]] = json[i][j];
        }
        rows.push(row);
      }
      let data = Object.keys(rows[0]);

      for (var i = 0; i < data.length; i++) {
        var container = document.createElement("div");
        var checkbox = document.createElement("input");

        checkbox.type = "checkbox";
        checkbox.name = data[i];
        checkbox.value = data[i];
        checkbox.id = "checkboxw" + i;

        var dropdown = document.createElement("select");
        dropdown.id = "dropdown" + i;
        dropdown.disabled = true;
        dropdown.hidden = true;
        for (let data in dropdownData) {
          let option = document.createElement("option");
          option.setAttribute("value", data);
          let optiontext = document.createTextNode(dropdownData[data]);
          option.appendChild(optiontext);
          dropdown.appendChild(option);
        }

        var drop2down = document.createElement("select");
        drop2down.id = "drop2down" + i;
        drop2down.disabled = true;
        drop2down.hidden = true;
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
    });
  }).catch(function (error) {
    console.log(error);
  });
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

export async function manipulateData() {
  var selectData = getType(selected);
  document.getElementById;
  manipulatedData = rows;
  try {
    await Excel.run(async (context) => {
      for (var i = 0; i < rows.length; i++) {
        for (var j = 0; j < selectData.length; j++) {
          if (selectData[j][1] == "string" && selectData[j][2] == "genRand") {
            manipulatedData[i][selectData[j][0]] = genRandString(7);
          }
          if (selectData[j][1] == "num" && selectData[j][2] == "genRand") {
            manipulatedData[i][selectData[j][0]] = genRandNum(5);
          }
          if (selectData[j][1] == "date" && selectData[j][2] == "genRand") {
            manipulatedData[i][selectData[j][0]] = genRandDate(new Date(2012, 0, 1), new Date());
          }
          if (selectData[j][2] == "encFile") {
            manipulatedData[i][selectData[j][0]] = md5(json[i][selectData[j][0]]);
          }
        }
      }
      let ws = context.workbook.worksheets.getActiveWorksheet();
      let range = ws.getUsedRange();
      range.load("values");
      await context.sync().then(function () {
        let data3 = [];
        for (var i = 0; i < rows.length; i++) {
          data3.push(Object.values(manipulatedData[i]));
        }
        data3.splice(0, 0, columns);
        range.values = data3;
        range.format.autofitColumns();
      });
    });
  } catch (error) {
    console.error(error);
  }
}
