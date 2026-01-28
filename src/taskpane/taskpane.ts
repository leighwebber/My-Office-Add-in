/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    document.getElementById("open-dialog").onclick = () => tryCatch(openDialog);
    document.getElementById("create-icon").onclick = () => tryCatch(insertIcon);
    document.getElementById("insert-stage-diagram").onclick = () => tryCatch(insertStageDiagram);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
  }
});
async function insertParagraph() {
    await Word.run(async (context) => {

        // TODO1: Queue commands to insert a paragraph into the document.
        const docBody = context.document.body;
        docBody.insertParagraph("The VERY SLOW brown fox jumps over the lazy dog.",
            Word.InsertLocation.start);
        await context.sync();
        
    });
}
async function insertIcon() {
    await Word.run(async (context) => {
        if(dialog){
            /* const messageObject = { messageType: "sillyStuff", text: "Hello there"};
            var jsonMessage = JSON.stringify(messageObject);
            dialog.messageChild(jsonMessage); */
            const iconObject = { messageType: "iconPlace", xPct: 0.30, yPct: 0.25, width: 40, height: 40};
            const jsonMessage = JSON.stringify(iconObject);
            dialog.messageChild(jsonMessage);
        }
    });
}
async function insertStageDiagram() {
    await Word.run(async (context) => {
        if(dialog){
            /* const messageObject = { messageType: "sillyStuff", text: "Hello there"};
            var jsonMessage = JSON.stringify(messageObject);
            dialog.messageChild(jsonMessage); */
            const messageObject = { messageType: "imageLoad", src: "../../assets/AnneFrankSet.jpg"};
            const jsonMessage = JSON.stringify(messageObject);
            dialog.messageChild(jsonMessage);
        }
    });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}
let dialog = null;
function openDialog() {
    Office.context.ui.displayDialogAsync(
    'https://localhost:3000/popup.html',
    {height: 45, width: 55},

      function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
  );
}

function processMessage(arg) {
    document.getElementById("user-name").innerHTML = arg.message;
    // dialog.close();
}
