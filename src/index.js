import Excel from 'exceljs/dist/es5/exceljs.browser';
import { saveAs } from 'file-saver'

const excelData = [];
const $dropZone = document.querySelector('.js-drop');

$dropZone.addEventListener("drop", function dropHandler(e) {
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

        // Read worksheet information from drag-and-dropped excel sheets
        const buffer = reader.result;
        const wb = new Excel.Workbook();
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

function transformRow(row) {
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
  if (excelData.length === 0) {
    return alert("No excel files were loaded");
  }

  const wb = new Excel.Workbook();
  const ws = wb.addWorksheet();
  excelData.forEach(sheet => {
    // this writes headers for all files
    sheet.eachRow(row => {
      ws.addRow(transformRow(row));
    });
  });
  wb.xlsx.writeBuffer()
    .then(buf => {
      saveAs(new Blob([buf]), 'abc.xlsx')
    });
});