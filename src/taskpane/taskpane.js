/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Office.js is ready in Outlook");

    const btn = document.querySelector("#btn");
    if (btn) {
      btn.addEventListener("click", () => run());
    } else {
      console.warn("Button not found!");
    }
  }
});

async function run() {
  const item = Office.context.mailbox.item;
  console.log(item);
}
