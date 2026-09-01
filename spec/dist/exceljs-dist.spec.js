const fs = require('fs');
const createMissingReferencesWorkbook = require('../utils/missing-references-workbook');

const exists = path => new Promise(resolve => fs.exists(path, resolve));

describe('ExcelJS', () => {
  describe('dist folder', () => {
    it('should include LICENSE', async () => {
      expect(await exists('./dist/LICENSE')).to.be.true()
    });
    it('should include exceljs.js', async () => {
      expect(await exists('./dist/exceljs.js')).to.be.true()
    });
    it('should include exceljs.min.js', async () => {
      expect(await exists('./dist/exceljs.min.js')).to.be.true()
    });
    it('should include exceljs.bare.js', async () => {
      expect(await exists('./dist/exceljs.bare.js')).to.be.true()
    });
    it('should include exceljs.bare.min.js', async () => {
      expect(await exists('./dist/exceljs.bare.min.js')).to.be.true()
    });
    it('should include es5/index', async () => {
      expect(await exists('./dist/es5/index.js')).to.be.true()
    });
  });

  describe('browser distribution', () => {
    for (const distribution of ['exceljs.js', 'exceljs.min.js']) {
      it(`should parse missing row and cell references from ${distribution}`, async () => {
        const ExcelJS = require(`../../dist/${distribution}`);
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await createMissingReferencesWorkbook(ExcelJS));

        expect(workbook.getWorksheet('First').getCell('B3').value).to.equal('Beta');
        expect(workbook.getWorksheet('Second').getCell('A2').value).to.equal('X');
      });
    }
  });
});
