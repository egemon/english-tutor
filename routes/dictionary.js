let express = require('express');
let router = express.Router();
let XLSX = require('xlsx');

/* GET users listing. */
router.post('/', function(req, res, next) {
  let xlsFile;

    if (!req.files) {
        res.send('No files were uploaded.');
        return;
    }

  let data = new Uint8Array(req.files.xlsFile.data);
  let arr = [];
  for(let i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
  let bstr = arr.join("");

  /* Call XLSX */
  let workbook = XLSX.read(bstr, {type:"binary"});
  let first_sheet_name = workbook.SheetNames[0];

  /* Get worksheet */
  let worksheet = workbook.Sheets[first_sheet_name];
  let dictionary = {};
  let rows = XLSX.utils.sheet_to_json(worksheet, {header:1});
  rows.forEach((row)=>{
    dictionary[row[2]] = row[3];
  });

  res.render('dictionary', {dictionary});
});

module.exports = router;
