const ExcelJS = verquire('exceljs');
const createMissingReferencesWorkbook = require('../../utils/missing-references-workbook');

describe('missing row and cell references', () => {
  it('loads workbooks whose worksheet coordinates are implicit', async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await createMissingReferencesWorkbook(ExcelJS));

    const firstWorksheet = workbook.getWorksheet('First');
    expect(firstWorksheet.getCell('A1').value).to.equal('id');
    expect(firstWorksheet.getCell('B1').value).to.equal('name');
    expect(firstWorksheet.getCell('A2').value).to.equal(1);
    expect(firstWorksheet.getCell('B3').value).to.equal('Beta');

    const secondWorksheet = workbook.getWorksheet('Second');
    expect(secondWorksheet.getCell('A1').value).to.equal('code');
    expect(secondWorksheet.getCell('A2').value).to.equal('X');
  });
});
