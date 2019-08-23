import Excel from 'exceljs/dist/es5/exceljs.browser';
import { saveAs } from 'file-saver'

// Set up drag-and-drop functionality
const $dropZone = document.querySelector('.js-drop');
console.log($dropZone);
$dropZone.addEventListener("drop", function dropHandler(e) {
  console.log("Something was dropped");
  e.preventDefault();

  if (e.dataTransfer.items) {
    for (let i = 0; i < e.dataTransfer.items.length; i++) {
      const currentItem = e.dataTransfer.items[i];
      const file = currentItem.getAsFile();
      console.log(`file: ${i}`, file);
    }
  }
});
$dropZone.addEventListener("dragover", function(e) {
  e.preventDefault();
});



// Generate a new Excel file for download
(async() => {
  const workbook = new Excel.Workbook();
  const sheet = workbook.addWorksheet('Test sheet', {
    properties: {
      tabColor: {
        argb:'FFC0000'
      }
    }
  });
  const row = sheet.addRow(['a', 'b', 'c']);
  row.font = { bold: true };

  const buffer = await workbook.xlsx.writeBuffer();
  // saveAs(new Blob([buffer]), 'abc.xlsx');
})();
