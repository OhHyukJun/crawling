const axios = require('axios');
const cheerio = require('cheerio');
const xlsx = require('xlsx');
const iconv = require('iconv-lite');

const url = 'http://addon.jinhakapply.com/RatioV1/RatioH/Ratio10950451.html';
// 진학사 url
//const url = 'http://ratio.uwayapply.com/Sl5KcldhVlc4TjlKJmB9YWhKemZUZg==';
// 유웨이 url
async function crawling() {
  try {
    const html = await axios.get(url,{
      responseType: 'arraybuffer'
    });//axios를 사용해 데이터를 불러온다

    const data = iconv.decode(html.data, 'utf-8').toString();
    // 유웨이 일때는 utf-8을 euc-kr로 변경
    const Data = cheerio.load(data);
    // cheerio를 사용하여 data를 파싱하고 Data 변수에 할당하여 사용
    const rows = [];
    // 데이터를 저장할 rows 배열
    Data('tr').each((index, element) => {
        const rowData = [];
        Data(element)
          .find('td')
          .toArray()
          .forEach(el => {
            rowData.push(Data(el).text());
          });
        rows.push(rowData);
    }); 
    /*
    html 코드에서 tr요소를 찾고 each를 사용해서 반복 
    rowData 배열에 td 요소의 값 저장
    td를 찾고 toArray를 사용해 배열로 변환 후 forEach를 사용하여 반복 후 rowData에 push
    Data(el).text() -> Data(el) = Cheerio 객체 생성
    마지막으로 td 값을 rows 배열에 추가함
    */
      
    const xlsxPath = './xlsx/crawling_data.xlsx';
    //xlsx 파일을 만들 경로
    writeToXLSX(rows, xlsxPath);
    //xlsx 파일 작성
    console.log('성공');
  } catch (error) {
    console.error('에러', error);
  }
}

function writeToXLSX(data, xlsxPath) {
  const workbook = xlsx.utils.book_new();
  //통합 문서 생성
  const worksheet = xlsx.utils.aoa_to_sheet(data);
  // 데이터를 작성할 워크시트 생성
  const range = xlsx.utils.decode_range(worksheet['!ref']);
  // xlsx.utils.decode_range는 워크시트 범위를 디코딩해준다 !ref가 범위
  worksheet['!ref'] = xlsx.utils.encode_range(range);
  // 디코딩된 range를 다시 인코딩하여 실제 크기를 넘는지 체크
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  // 첫 시트에 데이터 추가
  xlsx.writeFile(workbook, xlsxPath);
  // 액셀 파일 생성
}

crawling();
