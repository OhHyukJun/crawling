/*csv 란?
	Comma Separated Value라고 하며 콤마로 이루어진 파일이라고 볼 수 있습니다. 
	Csv 파일을 파싱할 때 해당 파일은 버퍼 파일(0,1로 이루어진 파일)이기 때문에 
    인코딩을 해야한다는 점을 주의해주세요.
*/			
/*
const parse = require('csv-parse/lib/index.js');
const fs = require('fs'); //파일 시스템 모듈을 불러온다
const csv = fs.reqdFileSync('csv/data.csv'); //파일을 불러오는 함수

csv.toString('utf-8') //encoding 해주기!

console.log('result: ', csv.toString('utf-8'));

const records = parse(csv.toString('utf-8'))
records.forEach((movieInfo, i) => {
	console.log('',movieInfo[0],'',movieInfo[1]);
});
*/
const axios = require('axios');
const cheerio = require('cheerio');
const fs = require('fs');
const csvWriter = require('csv-writer').createObjectCsvWriter;
const csv = require('csv-parser');
const XLSX = require('xlsx');
const iconv = require('iconv-lite'); //한글깨짐방지

const url = 'http://addon.jinhakapply.com/RatioV1/RatioH/Ratio10550281.html';

async function crawlWebsite() {
  try {
    await axios({
      url: url,
      method: "GET",
      responseType: "arraybuffer",
    }).then(async (html) => {
                const data = iconv.decode(html.data, "utf-8").toString();
    const $ = cheerio.load(data);
    const textContent = $('td.unit').text();
    
    //textContent.toString('EUC-KR');
    //textContent.encoding('euc-kr');
    
    const rows = [{ content: textContent }];
    
    const csvPath = 'website_content.csv';
    await writeToCSV(rows, csvPath);

    const workbook = XLSX.utils.book_new();
    const worksheet = await csvToWorksheet(csvPath);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    const xlsxPath = 'website_content.xlsx';
    XLSX.writeFile(workbook, xlsxPath);
  });
  } catch (error) {
    console.error('에러', error);
  }
}

async function writeToCSV(rows, csvPath) {
    return new Promise((resolve, reject) => {
      const csvWriterInstance = csvWriter({
        path: csvPath,
        header: [{ id: 'content', title: 'Content' }]
      });
      csvWriterInstance
        .writeRecords(rows)
        .then(() => resolve())
        .catch(reject);
    });
  }

  async function csvToWorksheet(csvPath) {
    return new Promise((resolve, reject) => {
      const data = [];
      fs.createReadStream(csvPath)
        .pipe(csv())
        .on('data', row => data.push(row))
        .on('end', () => {
          const worksheet = XLSX.utils.json_to_sheet(data);
          resolve(worksheet);
        })
        .on('error', reject);
    });
  }


crawlWebsite();

