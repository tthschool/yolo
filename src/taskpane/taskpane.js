/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
   let btn  = document.querySelector('#btn')
   btn.addEventListener('click' ,()=> {
    console.log("jdjdjd");
   })
  }
});

async function run() {
  const item = Office.context.mailbox.item;
  console.log("jello");
}
