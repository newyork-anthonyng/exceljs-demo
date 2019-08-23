import Excel from 'exceljs/dist/es5/exceljs.browser';
import { saveAs } from 'file-saver'

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
  saveAs(new Blob([buffer]), 'abc.xlsx');
})();
