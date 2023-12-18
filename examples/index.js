const fs = require('fs').promises;
const path = require('path');
const XlsxTemplater = require('../index')
async function main() {
  const template = path.join(__dirname, './resources/template.xlsx');
  const xlsxTemplater = new XlsxTemplater(template);
  await xlsxTemplater.parse();
  await xlsxTemplater.render({
    name: 'TEST',
    logo: path.join(__dirname, './resources/logo.png'),
    code: 'https://www.baidu.com',
    creator: 'xh4010',
    creatAt: new Date(),
    items: [{
        name: "Product A",
        quantity: 5,
        price: 10,
        desc:'',
      },
      {
        name: "Product B",
        quantity: 1,
        price:  20,
        desc:'note'
      },
    ],
  }, {})
  const pdfBuf = await xlsxTemplater.export();
  const filePath = path.join(__dirname, './resources/output/report1.pdf')
  await fs.writeFile(filePath, pdfBuf)
}

main().catch(function (err) {
  console.error('Error: ', err);
});