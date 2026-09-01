const JSZip = require('jszip');

module.exports = async ExcelJS => {
  const sourceWorkbook = new ExcelJS.Workbook();
  const firstWorksheet = sourceWorkbook.addWorksheet('First');
  firstWorksheet.addRows([
    ['id', 'name'],
    [1, 'Alpha'],
    [2, 'Beta'],
  ]);

  const secondWorksheet = sourceWorkbook.addWorksheet('Second');
  secondWorksheet.addRows([['code'], ['X']]);

  const zip = await JSZip.loadAsync(await sourceWorkbook.xlsx.writeBuffer());
  const worksheetEntries = Object.values(zip.files).filter(
    entry => !entry.dir && /^xl\/worksheets\/sheet\d+[.]xml$/.test(entry.name)
  );

  await Promise.all(
    worksheetEntries.map(async entry => {
      const worksheetXml = await entry.async('string');
      zip.file(
        entry.name,
        worksheetXml.replace(/<row r="[^"]*"/g, '<row').replace(/<c r="[^"]*"/g, '<c')
      );
    })
  );

  return zip.generateAsync({type: 'nodebuffer'});
};
