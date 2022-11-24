/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

  }
});

export async function run() {

  return Word.run(async (context) => {
    //Khai báo vv.

    let resultContainer = document.getElementById("result-container");
    let buttonRun = document.getElementById("run");
    let wait = document.createElement("p");
    let status = document.getElementById("status-text");
    let text = document.getElementById("keyword").value;
    let numOfResult = document.getElementById("numOfResult").value;
    let textLength = document.getElementById("numOfText").value;
    //Thay đổi trạng thái
    if (!checkValid(numOfResult)) {
      status.innerText = "Số kết quả phải lớn hơn 0";
      return;
    }
    if (!checkValid(textLength)) {
      status.innerText = "Độ dài văn bản phải lớn hơn 0";
      return;
    }
    status.innerText = "Đang xử lý...";
    wait.innerText = `Đang chờ ${numOfResult} kết quả .....`;

    buttonRun.disabled = true;
    resultContainer.appendChild(wait);
    // call API generation
    let result = await userAction(text, numOfResult, textLength);
    if (result === "error") {
      status.innerText = "Lỗi kết nối đến server";
      buttonRun.disabled = false;
      resultContainer.removeChild(wait);
      return;
    }
    resultContainer.removeChild(wait);
    status.innerText = "Đã xong";
    // in ra kết quả
    for (let i = 0; i < result.text.length; i++) {
      let resultTag = document.createElement("p");
      resultTag.innerText = `Kết quả ${i + 1}:\n${result.text[i]}`;
      resultContainer.appendChild(resultTag);
    }
    buttonRun.disabled = false;
    await context.sync();
  });
}
async function userAction(text, numOfResult, textLength) {
  const rawResponse = await fetch('http://localhost:8080/generate', {
    method: 'POST',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ text: text, numOfResult: parseInt(numOfResult), textLength: parseInt(textLength) })

  });
  if (rawResponse.ok) {
    return rawResponse.json();
  }
  else {
    return "error"
  }
}
function checkValid(input) {
  if (parseInt(input) < 0) {
    return false;
  }
  return true;
}

