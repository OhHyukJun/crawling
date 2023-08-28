const axios = require('axios');
const cheerio = require('cheerio');
const xlsx = require('xlsx');
const iconv = require('iconv-lite');

const url = 'http://addon.jinhakapply.com/RatioV1/RatioH/Ratio11640251.html';
// 진학사 url
//const url = 'http://ratio.uwayapply.com/Sl5KcldhVlc4TjlKJmB9YWhKemZUZg==';
// 유웨이 url
async function crawling() {
  try {
    const html = await axios.get(url,{
      responseType: 'arraybuffer'
    });//axios를 사용해 데이터를 불러온다

    const data = iconv.decode(html.data, 'utf-8').toString();

    //const data = iconv.decode(html.data, 'euc-kr').toString();
    // 유웨이 일때는 utf-8을 euc-kr로 변경

    const Data = cheerio.load(data);
    // cheerio를 사용하여 data를 파싱하고 Data 변수에 할당하여 사용

    const rows = [];
    // 데이터를 저장할 rows 배열

    // Data('tr').each((index, element) => {
    //     const rowData = [];
    //     Data(element)
    //       .find('td')
    //       .toArray()
    //       .forEach(el => {
    //         rowData.push(Data(el).text());
    //       });
    //     rows.push(rowData);
    // }); 
    /*
    html 코드에서 tr요소를 찾고 each를 사용해서 반복 
    rowData 배열에 td 요소의 값 저장
    td를 찾고 toArray를 사용해 배열로 변환 후 forEach를 사용하여 반복 후 rowData에 push
    Data(el).text() -> Data(el) = Cheerio 객체 생성
    마지막으로 td 값을 rows 배열에 추가함
    */
      
    const $ = cheerio.load(data);
    var ExcelData = [];
    $("div[id^='SelType'] table").each(function(){
        var Table = $(this), FirstTHCnt = Table.find("tbody tr:eq(0):has(th)").find("th").length;
        
        if (FirstTHCnt == 7){ // 계열 + 대학 + 최대모집 + 지원인원
          var SelType = ''; //Table.prev('h2').find("strong").text();
          var RowData = Table.find("tr:has(td):not(.total)").map(function(){
              var TR = $(this), TDGroup = TR.find("td"), TDCnt = TDGroup.length;
              // var Major = (TDCnt == 7) ? TDGroup.eq(2).text() : 
              //     (TDCnt == 4) ? TDGroup.eq(1).text() : 
              //     (TDCnt == 3) ? TDGroup.eq(0).text() : "";

              var Major = TDGroup.eq(2).text(), ApplyCnt = TDGroup.eq(5).text();

              // 가장 가까운 값이 있는 TR 찾기
              // var ClosestTRHasData = TR.prev("tr:has(td[rowSpan]:not([class]))")


              var Personnel = "", Ratio = "";
              if(TDCnt == 7){
                Major = TDGroup.eq(2).text();
                Personnel = TDGroup.eq(4).text();
                ApplyCnt = TDGroup.eq(5).text();
                console.log(TDGroup.eq(6));
                Ratio = TDGroup.eq(6);
              }
              else if(TDCnt == 4){
                Major = TDGroup.eq(1).text();
                ApplyCnt = TDGroup.eq(3).text();
              }
              else if(TDCnt == 3){
                Major = TDGroup.eq(0).text();
                ApplyCnt = TDGroup.eq(2).text();
                
              }

              return {SelType : SelType, Major : Major, Personnel : Personnel, ApplyCnt : ApplyCnt, Ratio : Ratio}
          }).get();

          ExcelData = ExcelData.concat(RowData);
        }
        else if(FirstTHCnt == 6){ // 계열 + 대학 + 최대모집 + 지원인원
          var SelType = ''; //Table.prev('h2').find("strong").text();
          var RowData = Table.find("tr:has(td):not(.total)").map(function(){
              var TR = $(this), TDGroup = TR.find("td"), TDCnt = TDGroup.length;
              // var Major = (TDCnt == 7) ? TDGroup.eq(2).text() : 
              //     (TDCnt == 4) ? TDGroup.eq(1).text() : 
              //     (TDCnt == 3) ? TDGroup.eq(0).text() : "";

              var Major = TDGroup.eq(0).text(), ApplyCnt = TDGroup.eq(1).text();

              // 가장 가까운 값이 있는 TR 찾기
              // var ClosestTRHasData = TR.prev("tr:has(td[rowSpan]:not([class]))")


              var Personnel = "", Ratio = "";
              if(TDCnt == 6){
                Major = TDGroup.eq(1).text();
                Personnel = TDGroup.eq(3).text();
                ApplyCnt = TDGroup.eq(4).text();
                Ratio = TDGroup.eq(5).text();
              }
              else if(TDCnt == 5){
                Major = TDGroup.eq(1).text();
                ApplyCnt = TDGroup.eq(3).text();
                Ratio = TDGroup.eq(4).text();
              }
              else if(TDCnt == 4){
                Major = TDGroup.eq(0).text();
                ApplyCnt = TDGroup.eq(2).text();
                Ratio = TDGroup.eq(3).text();
              }
              else if(TDCnt == 2){
                Major = TDGroup.eq(0).text();
                ApplyCnt = TDGroup.eq(1).text();
              }

              return {SelType : SelType, Major : Major, Personnel : Personnel, ApplyCnt : ApplyCnt, Ratio : Ratio}
          }).get();

          ExcelData = ExcelData.concat(RowData);
        }
        else if(FirstTHCnt == 5){ // 계열 + 대학 + 최대모집 + 지원인원
          var SelType = ''; //Table.prev('h2').find("strong").text();
          var RowData = Table.find("tr:has(td):not(.total)").map(function(){
              var TR = $(this), TDGroup = TR.find("td"), TDCnt = TDGroup.length;
              // var Major = (TDCnt == 7) ? TDGroup.eq(2).text() : 
              //     (TDCnt == 4) ? TDGroup.eq(1).text() : 
              //     (TDCnt == 3) ? TDGroup.eq(0).text() : "";

              var Major = TDGroup.eq(1).text(), ApplyCnt = TDGroup.eq(3).text();

              // 가장 가까운 값이 있는 TR 찾기
              // var ClosestTRHasData = TR.prev("tr:has(td[rowSpan]:not([class]))")


              var Personnel = "", Ratio = "";
              if(TDCnt == 5){
                Major = TDGroup.eq(1).text();
                Personnel = TDGroup.eq(2).text();
                ApplyCnt = TDGroup.eq(3).text();
                Ratio = TDGroup.eq(4).text();
              }
              else if(TDCnt == 4){
                Major = TDGroup.eq(0).text();
                Personnel = TDGroup.eq(1).text();
                ApplyCnt = TDGroup.eq(2).text();
                Ratio = TDGroup.eq(3).text();
              }
              else if(TDCnt == 2){
                Major = TDGroup.eq(0).text();
                ApplyCnt = TDGroup.eq(1).text();
              }

              return {SelType : SelType, Major : Major, Personnel : Personnel, ApplyCnt : ApplyCnt, Ratio : Ratio}
          }).get();

          // RowData.forEach((item, index, array) => {
          //   if(!item.Personnel) item.Personnel = array[index-1].Personnel;
          //   if(!item.Ratio) item.Ratio = array[index-1].Ratio;
          // });

          // Data.push(RowData); // 2차원 배열이 됨
          ExcelData = ExcelData.concat(RowData); // 배열끼리 병합
        }
        else if(FirstTHCnt == 3){ // 계열 + 대학 + 최대모집 + 지원인원
          var SelType = ''; //Table.prev('h2').find("strong").text();
          var RowData = Table.find("tr:has(td):not(.total)").map(function(){
              var TR = $(this), TDGroup = TR.find("td"), TDCnt = TDGroup.length;
              // var Major = (TDCnt == 7) ? TDGroup.eq(2).text() : 
              //     (TDCnt == 4) ? TDGroup.eq(1).text() : 
              //     (TDCnt == 3) ? TDGroup.eq(0).text() : "";

              var Major = TDGroup.eq(0).text(), ApplyCnt = TDGroup.eq(2).text();

              // 가장 가까운 값이 있는 TR 찾기
              // var ClosestTRHasData = TR.prev("tr:has(td[rowSpan]:not([class]))")


              //var Personnel = "", Ratio = "";
              if(TDCnt == 3){
                Major = TDGroup.eq(0).text();
                ApplyCnt = TDGroup.eq(2).text();
              
              }
              else if(TDCnt == 2){
                Major = TDGroup.eq(0).text();
                ApplyCnt = TDGroup.eq(1).text();
              }

              return {SelType : SelType, Major : Major, ApplyCnt : ApplyCnt,}
          }).get();

          ExcelData = ExcelData.concat(RowData);
        }
        else if(FirstTHCnt == 2){ // 계열 + 대학 + 최대모집 + 지원인원
          var SelType = ''; //Table.prev('h2').find("strong").text();
          var RowData = Table.find("tr:has(td):not(.total)").map(function(){
              var TR = $(this), TDGroup = TR.find("td"), TDCnt = TDGroup.length;
              // var Major = (TDCnt == 7) ? TDGroup.eq(2).text() : 
              //     (TDCnt == 4) ? TDGroup.eq(1).text() : 
              //     (TDCnt == 3) ? TDGroup.eq(0).text() : "";

              var Major = TDGroup.eq(0).text(), ApplyCnt = TDGroup.eq(1).text();

              // 가장 가까운 값이 있는 TR 찾기
              // var ClosestTRHasData = TR.prev("tr:has(td[rowSpan]:not([class]))")


              //var Personnel = "", Ratio = "";

              return {SelType : SelType, Major : Major, ApplyCnt : ApplyCnt,}
          }).get();

          ExcelData = ExcelData.concat(RowData);
        }
    });
    // console.log(ExcelData);

    let merge = [], MergeP = null, MergeR = null;

    ExcelData.forEach( (dr, i, arr) => {

      // { s: { r: 1, c: 2 }, e: { r: 10, c: 2 } }

      // Personnel 이 비어있을 떄 C는 2로 고정
      // Ratio 이 비어있을 떄 C는 4로 고정

      // 빈 값이 있을때 s가 등록됨 (바로 직전 row id)
      // 값이 있을 때 e가 등록됨 (바로 직전 row id)

      if(MergeP === null) {
        if(!dr.Personnel){
          MergeP = {s : {r : i, c : 2}};
        }
      }
      else if(i > 0 && dr.Personnel){
        if(MergeP['e'] === undefined) {
          MergeP['e'] = {r : i, c : 2};

          merge.push(MergeP);
          MergeP = null;
        }
      }
      else if(i == arr.length - 1){
        MergeP['e'] = {r : i + 1, c : 2};
        merge.push(MergeP);
        MergeP = null;
      }

      // if(MergeP['s'] === undefined) {
      //   MergeP['s'] = {r : i, c : 2};
      // }


      // if(!dr.Personnel){
      //   if(MergeP['s'] === undefined) {
      //     MergeP['s'] = {r : i, c : 2};
      //   }
      // }
      // else if(i > 0){
      //   if(MergeP['e'] === undefined) {
      //     MergeP['e'] = {r : i, c : 2};
      //   }
      // }




      // if(!dr.Personnel){
      //   if(MergeItem['s'] === undefined) {
      //     MergeItem['s'] = {r : i, c : 2};
      //   }
      // }
      // else if(i > 0){
      //   if(MergeItem['e'] === undefined) {
      //     console.log(dr, i);
      //     MergeItem['e'] = {r : i, c : 2};
      //   }
      // }

    });

    console.log(merge);
    //console.log("숙명여대 연세대 서울여대 크롤링 가능");

    
    const xlsxPath = './xlsx/crawling_data.xlsx';

    const workbook = xlsx.utils.book_new();
    //통합 문서 생성
    //const worksheet = xlsx.utils.aoa_to_sheet(data);
    // 데이터를 작성할 워크시트 생성 aoa_to_Sheet 방식
  
    const worksheet = xlsx.utils.json_to_sheet(ExcelData);
    // sheet 생성 - json_to_sheet 방식
  
    // const merge = [
    //   { s: { r: 1, c: 2 }, e: { r: 10, c: 2 } },
    //   { s: { r: 1, c: 4 }, e: { r: 10, c: 4 } },
    //   { s: { r: 11, c: 2 }, e: { r: 34, c: 2 } },
    //   { s: { r: 11, c: 4 }, e: { r: 34, c: 4 } },
    // ];
    // worksheet["!merges"] = merge;




    const range = xlsx.utils.decode_range(worksheet['!ref']);
    // xlsx.utils.decode_range는 워크시트 범위를 디코딩해준다 !ref가 범위
    worksheet['!ref'] = xlsx.utils.encode_range(range);
    // 디코딩된 range를 다시 인코딩하여 실제 크기를 넘는지 체크 구글링 하던 중에 추가한 코드 잘 이해가 가지는 않는다
  
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    // 첫 시트에 데이터 추가 -> 파일이 추가될 때마다 sheet 수를 늘려주면 좋을거 같다
    xlsx.writeFile(workbook, xlsxPath);
    // 액셀 파일 생성
    
    //xlsx 파일을 만들 경로

    // writeToXLSX(ExcelData, xlsxPath);
    //xlsx 파일 작성

    console.log('성공');
    


  } catch (error) {
    console.error('에러', error);
  }
}

/*
function writeToXLSX(data, xlsxPath) {
  const workbook = xlsx.utils.book_new();
  //통합 문서 생성
  //const worksheet = xlsx.utils.aoa_to_sheet(data);
  // 데이터를 작성할 워크시트 생성 aoa_to_Sheet 방식

  const worksheet = xlsx.utils.json_to_sheet(data);
  // sheet 생성 - json_to_sheet 방식

  const range = xlsx.utils.decode_range(worksheet['!ref']);
  // xlsx.utils.decode_range는 워크시트 범위를 디코딩해준다 !ref가 범위
  worksheet['!ref'] = xlsx.utils.encode_range(range);
  // 디코딩된 range를 다시 인코딩하여 실제 크기를 넘는지 체크 구글링 하던 중에 추가한 코드 잘 이해가 가지는 않는다

  xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  // 첫 시트에 데이터 추가 -> 파일이 추가될 때마다 sheet 수를 늘려주면 좋을거 같다
  xlsx.writeFile(workbook, xlsxPath);
  // 액셀 파일 생성
}
*/
crawling();