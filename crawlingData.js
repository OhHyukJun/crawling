const axios = require('axios');
const cheerio = require('cheerio');
const xlsx = require('xlsx');
const fs = require('fs');
const iconv = require('iconv-lite');

const https = require('https');
const agent = new https.Agent({
  rejectUnauthorized: false,
  secureProtocol: 'TLSv1_2_method'
});

async function crawling(service) {
  try {
    const html = await axios.get(service.RatioLink,{
      responseType: 'arraybuffer',
      httpsAgent: agent
    });

    const data = iconv.decode(html.data, 'utf-8').toString();
    const $ = cheerio.load(data);

    var ExcelData = [];
    var MergeData = [];

    $("div[id^='SelType'] table[class]").each(function(){
      var table = $(this), THS = table.find("tbody tr:eq(0):has(th)").find("th:not(.c)");
      var TableHeader = THS.map(function(){ return $(this).text().trim() }).get();
      var HeaderLen = TableHeader.length;
      var SelTypeName = table.prev('h2').find("strong").text().trim();
      
      // Your crawling logic here...
    
    }); // This closing bracket was missing.
  } catch (error) {
    console.error('에러', error.message);
  }
}

// Load the data from output.json
let UnivNamesOnly;
try {
  UnivNamesOnly= JSON.parse(fs.readFileSync('./output.json'));
} catch (err) {
  console.error(`Error reading file from disk: ${err}`);
}

async function main() {

  let crawlingResultsArray=[];
  
   for(let item of UnivNamesOnly){
     let resultOfCrawlingPage= await crawling(item);
     
     if(resultOfCrawlingPage){
         item.ExcelData=resultOfCrawlingPage.ExcelData;
         item.MergeData=resultOfCrawlingPage.MergeData;
         crawlingResultsArray.push(item); 
     }
   }

   const workbook2=xlsx.utils.book_new();
   const worksheet2=xlsx.utils.json_to_sheet(crawlingResultsArray);
   xlsx.utils.book_append_sheet(workbook2,worksheet2,'Sheet1');
   
   xlsx.writeFile(workbook2,'crawledOutput.xlsx');   
}

main();
