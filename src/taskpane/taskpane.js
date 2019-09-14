import lemur from "./../lemur_head.js";
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */

    // Queue a commmand to search the document.
    const searchResults = context.document.body.search('lemur', {matchCase: false})

    context.load(searchResults);

    return context.sync().then(function () {
      searchResults.items.forEach((result) => {
        // Queue a commmand to insert HTML in to the beginning of the body.
        result.insertInlinePictureFromBase64(lemur, Word.InsertLocation.replace);
          })
    });

    
  });
}