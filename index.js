const axios = require('axios');
const cheerio = require('cheerio');
const xlsx = require('xlsx');
const fs = require('fs');
const url = 'https://apply.jinhakapply.com/SmartRatio';

axios.get(url)
  .then((response)=>{
    if(response.status === 200){
      const html = response.data;
      const $ = cheerio.load(html);

      let BindData = JSON.parse($("input[name='hdnResultNow']").val());
      
      const regions = ['서울', '경기', '인천'];
      BindData.Rows = BindData.Rows.filter(row => {
        let regionIndex = BindData.Columns.indexOf('PresidentName');
        return regions.includes(row[regionIndex]);
      });

      const ExcelData = BindData.Rows.map(r => {
        let arr = BindData.Columns.map((c,i) => [c, r[i]]);
        return Object.fromEntries(arr);
      });

      const UnivNamesOnly= ExcelData.map(data => ({ UnivName: data.UnivName,CategoryDisplayName:data.CategoryDisplayName, StartDate:data.ApplyFromTime, EndDate:data.ApplyToTime, RatioLink: data.RatioLink }));

      
      const workbook = xlsx.utils.book_new();
      const worksheet = xlsx.utils.json_to_sheet(UnivNamesOnly);
      xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

	  xlsx.writeFile(workbook,'output.xlsx');
    }
  });

  axios.get(url)
  .then((response)=>{
    if(response.status === 200){
      const html = response.data;
      const $ = cheerio.load(html);

      let BindData = JSON.parse($("input[name='hdnResultNow']").val());
      
      const regions = ['서울', '경기', '인천'];
      BindData.Rows = BindData.Rows.filter(row => {
        let regionIndex = BindData.Columns.indexOf('PresidentName');
        return regions.includes(row[regionIndex]);
      });

      const ExcelData = BindData.Rows.map(r => {
        let arr = BindData.Columns.map((c,i) => [c, r[i]]);
        return Object.fromEntries(arr);
      });

	  const UnivNamesOnly= ExcelData.map(data => ({ UnivName: data.UnivName, UnivServiceID: data.UnivServiceID, StartDate:data.ApplyFromTime, EndDate:data.ApplyToTime, CategoryDisplayName: data.CategoryDisplayName, RatioLink: data.RatioLink }));

	  // Save all objects in one JSON file
	  fs.writeFileSync('output.json', JSON.stringify(UnivNamesOnly, null, 5));
    }
  });