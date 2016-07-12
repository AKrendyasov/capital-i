/**
 * Created by samsung np on 16.06.2016.
 */
/**
 * Created by samsung np on 12.06.2016.
 */
var Excel = require('exceljs');
var fs = require('fs');

var fileNameParser = function (Lmonth) {
    console.log("Парсинг имени файла")
   var typeOfInputFile =Lmonth.split('',1).join();
    var almostparsed = Lmonth.split('_');
    var fullyparsed = almostparsed[1].split('-');
    var pathToFile = __dirname + '/uploads/' + Lmonth;
/*    tryer(fullyparsed[0], fullyparsed[1], pathToFile,typeOfInputFile);*/
    dbConnector(fullyparsed[0], fullyparsed[1], pathToFile,typeOfInputFile)

};

var dbConnector = function(year, month,loadedFile,typeOfInputFile){
    console.log("Вход в функцию подключения к БД осуществлен");
    var dbBook = new Excel.Workbook();
    dbBook.xlsx.readFile(__dirname + '/db.xlsx')
        .then(function (dbook) {
            console.log("Подключение к БД осуществлено");
            if(typeOfInputFile=="c"){
                var worksheet = dbook.getWorksheet(2);
                console.log("Загружена корректировка")
            }
            else {
                var worksheet = dbook.getWorksheet(1);
                console.log("Загружена первичка")
            }
            recordCreator(year, month,loadedFile,worksheet,dbBook);


        });
};

var recordCreator=function (year, month,loadedFile,dbWorksheet,dbBook) {
    console.log("Вход в функцию создания записей для БД осуществлен");
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(loadedFile)
        .then(function () {
            var worksheet = workbook.getWorksheet('Лист1');
            var lastTableRow = worksheet.getRow(worksheet.lastRow.number-4);
            console.log("Номер последней строки в экселе"+lastTableRow.number);
            if(lastTableRow.number-11>0){
                for (i = 11; i < lastTableRow.number; i++) {
                    console.log("Номер обрабатываемой строки " + i);
                    var row=worksheet.getRow(i);
                    var controlCompany=worksheet.getCell('A6').value.split(':').splice(1, 1).join();
                    var finalizedRow=row.values;
                    finalizedRow.push(controlCompany);
                    var house = new House(year, month, finalizedRow);
                    var allowInput=true;
                    dbWorksheet.eachRow(function(row, rowNumber) {
                        if(house.company==row.values[1]&&house.adress==row.values[2]&&house.year==row.values[3]&&house.month==row.values[4]){
                            allowInput=false;
                            console.log("совпадение нашлось и запись " + allowInput)
                        }

                    });
                    if(allowInput==true){
                        console.log("Совпадение не нашлось и запись "+ allowInput);
                        console.log("Осуществлен вход в функцию создания новой записи в БД");

                        row = dbWorksheet.lastRow.number+1;
                        console.log("Номер строки в экселей"+row);
                        dbWorksheet.getCell('A'+row).value = house.company;
                        dbWorksheet.getCell('B'+row).value = house.adress;
                        dbWorksheet.getCell('C'+row).value = house.year;
                        dbWorksheet.getCell('D'+row).value = house.month;
                        dbWorksheet.getCell('E'+row).value = house.owner;
                        dbWorksheet.getCell('F'+row).value = house.ownersCount;
                        dbWorksheet.getCell('G'+row).value = house.square;
                        dbWorksheet.getCell('H'+row).value = house.lgotSquaree;
                        dbWorksheet.getCell('I'+row).value = house.techOmain;
                        dbWorksheet.getCell('J'+row).value = house.techOOther;
                        dbWorksheet.getCell('K'+row).value = house.naemSoc;
                        dbWorksheet.getCell('L'+row).value = house.naemKomm;
                        dbWorksheet.getCell('M'+row).value = house.naemBezDot;
                        dbWorksheet.getCell('N'+row).value = house.heat;
                        dbWorksheet.getCell('O'+row).value = house.capRepair;
                        dbWorksheet.getCell('P'+row).value = house.drainage;
                        dbWorksheet.getCell('Q'+row).value = house.waterSupply;
                        dbWorksheet.getCell('R'+row).value = house.eEnergy;
                        dbWorksheet.getCell('S'+row).value = house.compensation;
                        dbWorksheet.getCell('T'+row).value = house.changeSum;

                        console.log(dbWorksheet.getCell('B'+row).value);
                    }
                }
                dbBook.xlsx.writeFile(__dirname + '/db.xlsx')
                    .then(function() {
                        console.log("Запись окончена успешно");
                        // done
                    });
            } else {
                console.log("Файл некорректен")
            }
        });

};



function House(year, month,finalizedRow){
    this.adress=finalizedRow[2];
    this.year=year;
    this.month=month;
    this.owner="0";
    this.ownersCount="0";
    this.square="0";
    this.lgotSquaree="0";
    this.techOmain=finalizedRow[3];
    this.techOOther=finalizedRow[4];
    this.naemSoc=finalizedRow[5];
    this.naemKomm=finalizedRow[6];
    this.naemBezDot=finalizedRow[7];
    this.heat=finalizedRow[8];
    this.capRepair=finalizedRow[9];
    this.drainage=finalizedRow[10];
    this.waterSupply=finalizedRow[11];
    this.eEnergy=finalizedRow[12];
    this.compensation=finalizedRow[13];
    this.changeSum=finalizedRow[14];
    this.company=finalizedRow[15];

}

function newRecInputer(house,book,worksheet){
    console.log("Осуществлен вход в функцию создания новой записи в БД");

    var row = worksheet.lastRow.number+1;
    console.log(row);
    worksheet.getCell('A'+row).value = house.company;
    worksheet.getCell('B'+row).value = house.adress;
    worksheet.getCell('C'+row).value = house.year;
    worksheet.getCell('D'+row).value = house.month;
    worksheet.getCell('E'+row).value = house.owner;
    worksheet.getCell('F'+row).value = house.ownersCount;
    worksheet.getCell('G'+row).value = house.square;
    worksheet.getCell('H'+row).value = house.lgotSquaree;
    worksheet.getCell('I'+row).value = house.techOmain;
    worksheet.getCell('J'+row).value = house.techOOther;
    worksheet.getCell('K'+row).value = house.naemSoc;
    worksheet.getCell('L'+row).value = house.naemKomm;
    worksheet.getCell('M'+row).value = house.naemBezDot;
    worksheet.getCell('N'+row).value = house.heat;
    worksheet.getCell('O'+row).value = house.capRepair;
    worksheet.getCell('P'+row).value = house.drainage;
    worksheet.getCell('Q'+row).value = house.waterSupply;
    worksheet.getCell('R'+row).value = house.eEnergy;
    worksheet.getCell('S'+row).value = house.compensation;
    worksheet.getCell('T'+row).value = house.changeSum;

    console.log(worksheet.getCell('B'+row).value);
/*    book.xlsx.writeFile(__dirname + '/db.xlsx')
        .then(function() {
            console.log("Запись окончена успешно");
            // done
        });*/
}

/*

    var number="data"+n;

    var useMonth=Lmonth+".xlsx";
    console.log(useMonth)
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(useMonth)
        .then(function () {
            var worksheet = workbook.getWorksheet('My Sheet');
            var row = worksheet.lastRow;
            var finalizedRow=row.values;

            if (n==1)            {

                console.log(fs.writeFileSync('t.json', JSON.stringify(data, null, "\t")));
                fs.writeFileSync('t.json', JSON.stringify(data, null, "\t"));
            } else {};
            n = fs.readFileSync('t.json',"utf-8")
            data=JSON.parse(n);
            data[number].adress=workbook.getWorksheet('My Sheet').getCell('C4').value;
            data[number].year=workbook.getWorksheet('My Sheet').getCell('B4').value;
            data[number].month=workbook.getWorksheet('My Sheet').getCell('A4').value;
            data[number].owner=finalizedRow[1];
            data[number].ownersCount=finalizedRow[2];
            data[number].square=finalizedRow[3];
            data[number].lgotSquare=finalizedRow[4];
            data[number].techOmain=finalizedRow[5];
            data[number].techOOther=finalizedRow[6];
            data[number].naemSoc=finalizedRow[7];
            data[number].naemKomm=finalizedRow[8];
            data[number].naemBezDot=finalizedRow[9];
            data[number].heat=finalizedRow[10];
            data[number].capRepair=finalizedRow[11];
            data[number].drainage=finalizedRow[12];
            data[number].waterSupply=finalizedRow[13];
            data[number].eEnergy=finalizedRow[14];
            data[number].compensation=finalizedRow[15];

            // use workbook
            console.log( data[number].ownersCount);
            fs.writeFileSync('t.json', JSON.stringify(data, null, "\t"));

        })*/



exports.fileNameParser=fileNameParser;

