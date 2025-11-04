const XLSX = require('xlsx');

class Data {
  readExcel(filePath, sheetName) {
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet); // <-- returns array
  }
  writeExcel(filePath, sheetName, data) {
    const XLSX = require('xlsx');
    let workbook;
    try {
        workbook = XLSX.readFile(filePath);
        // Remove the old sheet if it exists
        if (workbook.SheetNames.includes(sheetName)) {
            delete workbook.Sheets[sheetName];
            const idx = workbook.SheetNames.indexOf(sheetName);
            if (idx > -1) workbook.SheetNames.splice(idx, 1);
        }
    } catch (e) {
        workbook = XLSX.utils.book_new();
    }
    const worksheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, filePath);
}
updateExcel(data,rowInput,key,value){
  const row = data.find(row => row.SerialNo === rowInput);
      switch (key) {
        case "FirstName":row.FirstName = value;
          break;
        case "LastName":row.LastName = value;
          break;
        case "Email":row.Email = value;
          break;
        case "Country":row.Country = value;
          break;
        case "Currency":row.Currency = value;
          break;
        case "Phone":row.Phone = value;
          break;
        case "Filter":row.Filter = value;
          break;
        case "Budget":row.Budget = value;
          break;
        case "Room":row.Room = value;
          break;
        case "Member":row.Member = value;
          break;
        case "Destination":row.Destination = value;
          break;
        case "Arrival":row.Arrival = value;
          break;
        case "Departure":row.Departure = value;
          break;
        case "Rules":row.Rules = value;
          break;
        case "Details":row.Details = value;
          break;
        case "PNR":row.PNR = value;
          break;
        case "PIN":row.PIN = value;
          break;
        case "Status":row.Status = value;
          break;
        case "Address":row.Address = value;
          break;
        default:break;
      } 
}
    }


module.exports = { Data };
        
