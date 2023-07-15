const axios = require('axios');
const cheerio = require('cheerio');
const fs = require('fs');
const xlsx = require('xlsx');
const iconv = require('iconv-lite');

const url = 'http://addon.jinhakapply.com/RatioV1/RatioH/Ratio10950451.html';

async function crawlWebsite() {
  try {
    const html = await axios({
      url: url,
      method: 'GET',
      responseType: 'arraybuffer',
    });

    const data = iconv.decode(html.data, 'utf-8').toString();
    const $ = cheerio.load(data);

    const rows = [];
    $('tr').each((index, element) => {
      const rowData = [];
      $(element)
        .find('td')
        .each((i, el) => {
          rowData.push($(el).text());
        });
      rows.push(rowData);
    });

    const xlsxPath = 'website_content.xlsx';
    writeToXLSX(rows, xlsxPath);
    console.log('성공');
  } catch (error) {
    console.error('에러', error);
  }
}

function writeToXLSX(data, xlsxPath) {
  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.aoa_to_sheet(data);
  const range = xlsx.utils.decode_range(worksheet['!ref']);
  worksheet['!ref'] = xlsx.utils.encode_range(range);
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  xlsx.writeFile(workbook, xlsxPath);
}

crawlWebsite();
