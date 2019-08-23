import Excel from 'exceljs/dist/es5/exceljs.browser';
import { saveAs } from 'file-saver'

// Set up drag-and-drop functionality
const excelData = [];

const $dropZone = document.querySelector('.js-drop');
console.log($dropZone);
$dropZone.addEventListener("drop", function dropHandler(e) {
  console.log("Something was dropped");
  e.preventDefault();

  if (e.dataTransfer.items) {
    for (let i = 0; i < e.dataTransfer.items.length; i++) {
      const currentItem = e.dataTransfer.items[i];
      const file = currentItem.getAsFile();
      console.log(file);
      const reader = new FileReader();
      reader.readAsArrayBuffer(file);
      reader.onload = () => {
        addUploadedFile(file);

        const buffer = reader.result;
        const wb = new Excel.Workbook();
        // Read excel sheet buffer
        wb.xlsx.load(buffer).then(workbook => {
          workbook.eachSheet((sheet, id) => {
            excelData.push(sheet);
          });

        });
      };
    }
  }
});

$dropZone.addEventListener("dragover", function(e) {
  e.preventDefault();
});

const $uploadedFiles = document.querySelector('.js-uploaded');
function addUploadedFile(file) {
  const $listItem = document.createElement('li');
  $listItem.innerText = `${file.name}`;
  $uploadedFiles.appendChild($listItem);
}

function transformRow(row, rowIndex) {
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

  return [
    orderDate, region, rep, units, unitCost, item, total
  ];
};

// Generate output excel file
const $generate = document.querySelector('.js-generate');
$generate.addEventListener('click', function() {
  console.log('%cgenerate button was clicked', 'background-color: green; color: white; font-size: 24px;');
  if (excelData.length === 0) {
    return alert("No excel files were loaded");
  }
  console.log(excelData);
  // workbook.xlsx.writeBuffer().then(buffer => {
  //   saveAs(new Blob([buffer]), 'abc.xlsx');
  // });
});