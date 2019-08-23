import Excel from 'exceljs/dist/es5/exceljs.browser';
import { saveAs } from 'file-saver'

const excelData = [];
const $dropZone = document.querySelector('.js-drop');
const $fileInput = document.querySelector('.js-file-input');
$fileInput.addEventListener('change', function(e) {
  const files = e.target.files;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];

    createExcelWorksheetFromFile(file);
  }
});

function createExcelWorksheetFromFile(file) {
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);

    reader.onload = () => {
      addUploadedFile(file); // handles UI updates

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

$dropZone.addEventListener("drop", function dropHandler(e) {
  e.preventDefault();

  if (e.dataTransfer.items) {
    for (let i = 0; i < e.dataTransfer.items.length; i++) {
      const currentItem = e.dataTransfer.items[i];
      const file = currentItem.getAsFile();
      createExcelWorksheetFromFile(file);
    }
  }
});

$dropZone.addEventListener("dragover", function(e) {
  e.preventDefault();
  $dropZone.classList.add("is-dragover");
});

['dragleave', 'dragend', 'drop'].forEach(eventType => {
  $dropZone.addEventListener(eventType, () => {
    $dropZone.classList.remove("is-dragover");
  });
});


const $uploadedFiles = document.querySelector('.js-uploaded');
function addUploadedFile(file) {
  const $listItem = document.createElement('li');
  $listItem.innerText = `${file.name}`;
  $uploadedFiles.appendChild($listItem);
}

const mapping = {
  date: 1,
  units: 5,
  unitCost: 6,
  total: 7
};
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
    row.values[mapping.date],
    row.values[mapping.units],
    row.values[mapping.unitCost],
    row.values[mapping.total]
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
  ws.addRow(['Date', 'Units', 'Unit cost', 'Total']);
  excelData.forEach(sheet => {
    sheet.eachRow((row, rowIndex) => {
      if (rowIndex === 1) {
        // skip header row
        return;
      }
      ws.addRow(transformRow(row));
    });
  });
  wb.xlsx.writeBuffer()
    .then(buf => {
      saveAs(new Blob([buf]), 'abc.xlsx')
    });
});