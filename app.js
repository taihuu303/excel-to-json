const excelToJson = require('convert-excel-to-json');
const fs = require('fs');
const path = require('path');

const URL_RESPONSE_JSON = path.resolve(__dirname, 'resources.json');

const isExistedFile = (path) => {
  if (fs.existsSync(path)) {
    return true;
  }
  return false;
}

const handleReadFile = () => {
  if (!isExistedFile(URL_RESPONSE_JSON)) return {};

  const getFile = fs.readFileSync(URL_RESPONSE_JSON);
  const rawData = !!getFile && JSON.parse(getFile);

  return rawData;
}

const handleWriteFile = (data) => {
  return fs.writeFileSync(URL_RESPONSE_JSON, JSON.stringify(data));
}

const nameSheet = 'SheetTest';
const getJsonData = excelToJson({
  sourceFile: 'test.xlsx',
  header: {
    rows: 2
  },
  sheets: [nameSheet],
  columnToKey: {
    'B': 'key',
    'C': 'id',
    'E': 'value'
  }
});
console.log(`\x1b[32mTaiTH -> getJsonData:`, getJsonData);

let result = {};

getJsonData[nameSheet].forEach(item => {
  const key = item.key + '_' + item.id;
  result[key] = item.value;
})

handleWriteFile(result);

console.log("TaiTH ~ Done");
