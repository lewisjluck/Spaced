/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { PDFDocument } from 'pdf-lib'
import { convertPdfToPng } from 'convert-pdf-png';

Office.onReady((info) => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("insert").onclick = run;
  }
});


export async function run() {
  /**
   * Insert your OneNote code here
   */
   try {
    await OneNote.run(async (context) => {
      var pdfjsLib =  await import("pdfjs-dist");
      var pdfjsWorker =  await import('pdfjs-dist/build/pdf.worker.entry')
      pdfjsLib.workerSrc = pdfjsWorker;
        var file = document.getElementById('file-upload').files[0];
        
        let fileReader = new FileReader();
        fileReader.onload = function(ev) {
          pdfjsLib.getDocument(fileReader.result).then(function getPdfHelloWorld(pdf) {
            //
            // Fetch the first page
            //
            console.log("HELLO")
            pdf.getPage(1).then(function getPageHelloWorld(page) {
              var scale = 1.5;
              var viewport = page.getViewport(scale);

              //
              // Prepare canvas using PDF page dimensions
              //
              var canvas = document.getElementById('convert-canvas');
              var context = canvas.getContext('2d');
              canvas.height = viewport.height;
              canvas.width = viewport.width;

              //
              // Render PDF page into canvas context
              //
              var task = page.render({canvasContext: context, viewport: viewport})
              task.promise.then(function(){
                const oneNotePage = context.application.getActivePage();
                var addedOutline = page.addOutline(40,90);
                addedOutline.appendHtml("<img src='{canvas.toDataURL('image/jpeg')}'></img>");
                console.log(canvas.toDataURL('image/jpeg'));
              });
            });
          });
          
          
        }
        fileReader.readAsArrayBuffer(file);


        
        /*
        convertPdfToPng(file, {
          outputType: 'callback',
          callback: callback
        });
    
        const callback = images => {
            // the function returns an array
            // every img is a normal file object
            images.forEach(img => {
              const url = URL.createObjectURL(img);
              console.log(url);
              

              
            });
        }*/
        
        // Get the current page content
        
        
    });
    } catch (error) {
        console.log("Error: " + error);
    }

}