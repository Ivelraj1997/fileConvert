var XLSX = require('xlsx');

let fileConverter=async ()=>{
    try {
        const workbook = XLSX.readFile('EWNworkstreamAutomationInput.xlsx');
        const sheet_name_list = workbook.SheetNames;
        let sheetValues=XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

        let wb = XLSX.utils.book_new()
        for(let i=0;i<sheetValues.length;i++){
            if(sheetValues[i]['Required Tasks']){
               let reqValue= sheetValues[i]['Required Tasks'];
               let formatedValue= reqValue.replace(/,/i, ' AND ').replace('|', ' OR ').replace(/&/i, ' AND ');
               sheetValues[i]['Rules']=formatedValue;
            }
        }
            let ws = XLSX.utils.json_to_sheet(sheetValues);
            XLSX.utils.book_append_sheet(wb, ws)
            let exportFileName = `EWNworkstreamAutomationOutput.xlsx`;
            XLSX.writeFile(wb, exportFileName);
      } catch (err) {
        console.log(err);
      }
}

fileConverter()