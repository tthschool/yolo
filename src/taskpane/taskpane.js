/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    console.log("dajskndlasjkndlkasndlk")
    const btn = document.querySelector("#btn");
    // btn.addEventListener("click", () => {
    //   console.log("clicked");
          document.getElementById("run").onclick = run;
    });


  }
});


async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  console.log("jello");
  
}
    console.log("dajskndlasjkndlkasndlk")
