/**
 * Created by samsung np on 06.07.2016.
 */
var Excel = require('exceljs');
var fs = require('fs');

function inputSelector (company,sum,month,year){

    console.log("Осуществлен вход в функцию ввода данных с формы");

    var payBook = new Excel.Workbook();
    payBook.xlsx.readFile(__dirname + '/pay.xlsx')
        .then(function (dbook) {
            var worksheet = dbook.getWorksheet(1);
            var usedRowNumber=worksheet.lastRow.number+1;
            worksheet.eachRow(function (row, rowNumber) {
                console.log(row.values[1]);
                console.log(row.values[2]);
                console.log(row.values[3]);
                console.log(row.values[4]);
                if(company==row.values[1]&&year==row.values[2]&&month==row.values[3]){
                    usedRowNumber=rowNumber
                }
            });
            console.log(usedRowNumber);
            console.log(company);
            console.log(sum);
            console.log(month);
            console.log(year);
            worksheet.getCell('A'+(usedRowNumber)).value = company;
            worksheet.getCell('B'+(usedRowNumber)).value = year;
            worksheet.getCell('C'+(usedRowNumber)).value = month;
            worksheet.getCell('D'+(usedRowNumber)).value = Number(sum);
            dbook.xlsx.writeFile(__dirname + '/pay.xlsx')
                .then(function() {
                    console.log("Запись выполнена");
                    // done
                });

        });
};


exports.inputSelector =inputSelector;