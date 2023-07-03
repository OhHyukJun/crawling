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

const url = 'http://addon.jinhakapply.com/RatioV1/RatioH/Ratio10190351.html';

async function crawlWebsite() {
  try {
    const response = await axios.get(url);
    const $ = cheerio.load(response.data);
    const textContent = $('td.rate4').text();
    textContent.toString('utf-8');
    // Convert text content to an array of objects for CSV writing
    const rows = [{ content: textContent }];

    const csvPath = 'website_content.csv';
    await writeToCSV(rows, csvPath);

    const workbook = XLSX.utils.book_new();
    const worksheet = await csvToWorksheet(csvPath);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    const xlsxPath = 'website_content.xlsx';
    XLSX.writeFile(workbook, xlsxPath);
    console.log('XLSX file created successfully.');
  } catch (error) {
    console.error('Error occurred during crawling:', error);
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
    const readStream = fs.createReadStream(csvPath);
    const worksheet = {};
    const data = [];
    readStream
      .pipe(csv())
      .on('data', row => data.push(row))
      .on('end', () => {
        worksheet['!ref'] = `A1:A${data.length}`;
        worksheet['A1'] = { t: 's', v: 'content' };
        data.forEach((row, index) => {
          worksheet[`A${index + 2}`] = { t: 's', v: row.content };
        });
        resolve(worksheet);
      })
      .on('error', reject);
  });
}

crawlWebsite();

