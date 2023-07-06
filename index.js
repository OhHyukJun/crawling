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
//const fs = require('fs');
const csvWriter = require('csv-writer').createObjectCsvWriter;
//const csv = require('csv-parser');
//const XLSX = require('xlsx');
const iconv = require('iconv-lite'); //한글깨짐방지

const url = 'http://addon.jinhakapply.com/RatioV1/RatioH/Ratio10080141.html';

async function crawlWebsite() {
  try {
    await axios({
      url: url,
      method: "GET",
      responseType: "arraybuffer",
    }).then(async (html) => {
                const data = iconv.decode(html.data, "utf-8").toString();
                //한글이 깨지는 것을 방지하기 위한 코드
    const $ = cheerio.load(data);
    const textContent = $('td.unit').text();
    //hrml 태그를 확인해 크롤링할  정보의 class 명을 사용하여 크롤링 진행
    
   
    
    const rows = [{ content: textContent }];
    //데이터를 rows로 삽입 -> 행렬로 삽입하고 싶다
    const csvPath = 'website_content.csv';
    //csv 생성 경로
    await writeToCSV(rows, csvPath)
      .then(() => console.log("성공"))
      .catch((error) => console.log('실패'));
      // 이거까진 잘되는 거 같음
  });
  } catch (error) {
    console.error('에러', error);
  }
}

async function writeToCSV(rows, csvPath) {
  try {
    const csvWriterInstance = csvWriter({
      path: csvPath,
      header: [{ id: 'content', title: '경쟁률 정보' }]
    });

    await csvWriterInstance.writeRecords(rows);
    //주어진 rows 배열의 데이터를 csv 파일에 작성
    
    return;
    //변환 성공값을 출력
  } catch (error) {
    throw error;
  }
}

crawlWebsite();

