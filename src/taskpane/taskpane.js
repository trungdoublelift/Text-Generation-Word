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
    let numOfResults = document.getElementById("numOfResults").value;
    let resultContainer = document.getElementById("result-container");
    let buttonRun=document.getElementById("run");
    let wait = document.createElement("p");
    //Thay đổi trạng thái
    wait.innerText = `Đang chờ ${numOfResults} kết quả .....`;
    buttonRun.disabled=true;
    resultContainer.appendChild(wait);
    // call API generation
    let result = await userAction();
    resultContainer.removeChild(wait);
    // in ra kết quả
    for (let i = 0; i < numOfResults; i++) {
      let resultTag = document.createElement("p");
      resultTag.innerText =`Kết quả ${i+1}:${result.text}`;
      resultContainer.appendChild(resultTag);
    }
    buttonRun.disabled=false;
    await context.sync();
  });
}
async function userAction() {
  const rawResponse = await fetch('http://localhost:8080/generate', {
    method: 'POST',
    headers: {
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ text: document.getElementById("keyword").value }),
  });
  if (rawResponse.ok) {
    return rawResponse.json();
  }
  else {
    return "error"
  }
}