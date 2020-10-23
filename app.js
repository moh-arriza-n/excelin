//variabel express-------------------------------------------------
var express = require('express');
var app = express();
var bodyParser = require('body-parser');

//variable urlencdparser-------------------------------------------
var urlencodedParser = bodyParser.urlencoded({extended:false})

//variabel exceljs
var Excel = require('exceljs');
var { response } = require('express');
var workbook = new Excel.Workbook();

//-----------------------------------------------------------------
app.use(express.static('public'));
app.get('/',function(req,res){
  res.sendFile(__dirname +"/"+"form.htm");
})
 
app.post('/hasil_output_JSON', urlencodedParser, (req, res) => {
        // Prepare output in JSON format

        response = {
            cell_G3: req.body.cell_G3,
            cell_G4: req.body.cell_G4,
            cell_G5: req.body.cell_G5,
            cell_G6: req.body.cell_G6,
            cell_K3: req.body.cell_D3,
            cell_K4: req.body.cell_D4,
            cell_K5: req.body.cell_D5,
            cell_K6: req.body.cell_D6,
            cell_M3: req.body.cell_M3,
            cell_M4: req.body.cell_M4,
            cell_M5: req.body.cell_M5,
            cell_M6: req.body.cell_M6,
            cell_T2: req.body.cell_T2

        }

        workbook.xlsx.readFile('test.xlsx')
        .then(function () {
            var worksheet = workbook.getWorksheet('2019 IO');
    
            worksheet.getCell('G3').value = parseInt(req.body.cell_G3);
            worksheet.getCell('G4').value = parseInt(req.body.cell_G4);
            worksheet.getCell('G5').value = parseInt(req.body.cell_G5);
            worksheet.getCell('G6').value = parseInt(req.body.cell_G6);
            worksheet.getCell('D3').value = parseInt(req.body.cell_D3);
            worksheet.getCell('D4').value = parseInt(req.body.cell_D4);
            worksheet.getCell('D5').value = parseInt(req.body.cell_D5);
            worksheet.getCell('D6').value = parseInt(req.body.cell_M6);
            worksheet.getCell('M3').value = parseInt(req.body.cell_M3);
            worksheet.getCell('M4').value = parseInt(req.body.cell_M4);
            worksheet.getCell('M5').value = parseInt(req.body.cell_M5);
            worksheet.getCell('M6').value = parseInt(req.body.cell_M6);
            worksheet.getCell('T2').value = parseInt(req.body.cell_T2);
        return workbook.xlsx.writeFile('test.xlsx');
        });


        console.log(response);
        res.end(JSON.stringify({response}));

        //dibawah ini untuk konversi string to integer , tapi bosok !
        //gunakan parseInt() pada value yang dituju dan target yang dituju.
        //res.send(JSON.stringify(response, function(key, value) { return parseInt(value) || value}));
    });



//------------------------------------------------------------------------

app.listen(8080,"localhost");
console.log("Warung sudah buka di http://localhost:8080");

//-----------------------------------------------------------------------
//WARNING
//
//Belum masalah emptyfield pada input element yang terus masuk ke excel yang dianggap nilainya null / kosong.


