/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    document.getElementById("get-selection-text").onclick = () => tryCatch(getSelectionText);
    document.getElementById("open-dialog").onclick = () => tryCatch(openDialog);
    document.getElementById("create-icon").onclick = () => tryCatch(insertIcon);
    document.getElementById ("insert-stage-diagram").onclick = () => tryCatch(insertStageDiagram);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
  }
});
let dialog = null;
import { processMessage, openDialog, tryCatch, insertStageDiagram, insertIcon,
    insertParagraph, getSelectionText } from './taskpaneFunctions'
