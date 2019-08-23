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
      const reader = new FileReader();
      reader.readAsArrayBuffer(file);
      reader.onload = () => {
        const buffer = reader.result;
        const wb = new Excel.Workbook();
        wb.xlsx.load(buffer).then(workbook => {
          console.log(workbook, 'workbook instance');
          workbook.eachSheet((sheet, id) => {
            sheet.eachRow(function transformRow(row, rowIndex) {
              const [
                _,
                orderDate,
                region,
                rep,
                item,
                units,
                unitCost,
                total
              ] = row.values;

              row.values = [
                orderDate, region, rep, units, unitCost, item, total
              ];
            });
          });
          // workbook.eachSheet((sheet, id) => {
          //   sheet.eachRow((row, rowIndex) => {
          //     console.log(row.values, rowIndex);
          //   })
          // });
          // const sheet = workbook.addWorksheet('Test sheet', {
          //   properties: {
          //     tabColor: {
          //       argb:'FFC0000'
          //     }
          //   }
          // });
          // const row = sheet.addRow(['a', 'b', 'c']);
          // row.font = { bold: true };

          workbook.xlsx.writeBuffer().then(buffer => {
            saveAs(new Blob([buffer]), 'abc.xlsx');
          });
        });
      };
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
