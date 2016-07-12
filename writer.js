/**
 * Created by samsung np on 16.06.2016.
 */
/**
 * Created by samsung np on 14.06.2016.
 */
var Excel = require('exceljs');
var fs = require('fs');
var fse = require('fs');

var reportSelector = function (usedReport,company,houseList,usedYears,usedMonths,res) {
    if (usedReport == "lostIncome") {
        lostIncome(company, houseList, usedYears, usedMonths, res);
        console.log("Запущен отчет " + usedReport);
    } else {
        lostIncomeCompany(company, houseList, usedYears, usedMonths, res);
        /*        console.log("Запрашиваемый отчет " +usedReport+ " не найден");*/
        console.log("Запущен отчет " + usedReport);
    }

};
function lostIncomeCompany(company,houseList,usedYears,usedMonths,res){
    dbinputerCompany(company,houseList,usedYears,usedMonths,res)
};

function dbinputerCompany(company,houseList,usedYears,usedMonths,res){
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(__dirname + '/db.xlsx')
        .then(function (db) {
            var sumTry=0;
            var bdArray = [];
            console.log('соединение с БД установлено');
            var worksheet = db.getWorksheet(1);
            var worksheet2 = db.getWorksheet(2);

            usedYears.forEach(function(item, i, arr) {
                usedMonths.forEach(function(mItem, mI, mArr) {
                    var primaryData= [company,"",item,mItem,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0];
                    var correctionData= [company,"",item,mItem,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0];
                    var totalData= [company,"",item,mItem,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0];
                    worksheet.eachRow(function (row, rowNumber) {
                        if(item==row.values[3]&&mItem==row.values[4]&&company==row.values[1]){
                            primaryData[0]=row.values[1];
                            primaryData[1]="";
                            primaryData[2]=row.values[3];
                            primaryData[3]=row.values[4];
                            primaryData[4]=primaryData[4]+row.values[5];
                            primaryData[5]=primaryData[5]+row.values[6];
                            primaryData[6]=primaryData[6]+row.values[7];
                            primaryData[7]=primaryData[7]+row.values[8];
                            primaryData[8]=primaryData[8]+row.values[9];
                            primaryData[9]=primaryData[9]+row.values[10];
                            primaryData[10]=primaryData[10]+row.values[11];
                            primaryData[11]=primaryData[11]+row.values[12];
                            primaryData[12]=primaryData[12]+row.values[13];
                            primaryData[13]=primaryData[13]+row.values[14];
                            primaryData[14]=primaryData[14]+row.values[15];
                            primaryData[15]=primaryData[15]+row.values[16];
                            primaryData[16]=primaryData[16]+row.values[17];
                            primaryData[17]=primaryData[17]+row.values[18];
                            primaryData[18]=primaryData[18]+row.values[19];
                            primaryData[19]=primaryData[19]+row.values[20];

                            console.log("Cовпадение найдено "+row.values[2] + "и значение"+row.values[19]);
                        }
                    });

                    worksheet2.eachRow(function (row, rowNumber) {
                        if(item==row.values[3]&&mItem==row.values[4]&&company==row.values[1]){
                            correctionData[0]=row.values[1];
                            correctionData[1]="";
                            correctionData[2]=row.values[3];
                            correctionData[3]=row.values[4];
                            correctionData[4]=correctionData[4]+row.values[5];
                            correctionData[5]=correctionData[5]+row.values[6];
                            correctionData[6]=correctionData[6]+row.values[7];
                            correctionData[7]=correctionData[7]+row.values[8];
                            correctionData[8]=correctionData[8]+row.values[9];
                            correctionData[9]=correctionData[9]+row.values[10];
                            correctionData[10]=correctionData[10]+row.values[11];
                            correctionData[11]=correctionData[11]+row.values[12];
                            correctionData[12]=correctionData[12]+row.values[13];
                            correctionData[13]=correctionData[13]+row.values[14];
                            correctionData[14]=correctionData[14]+row.values[15];
                            correctionData[15]=correctionData[15]+row.values[16];
                            correctionData[16]=correctionData[16]+row.values[17];
                            correctionData[17]=correctionData[17]+row.values[18];
                            correctionData[18]=correctionData[18]+row.values[19];
                            correctionData[19]=correctionData[19]+row.values[20];
                        }
                    });
                    console.log(primaryData);
                    console.log(correctionData);
                    for (i=4; i < correctionData.length; i++) {
                        totalData[i]=(Number(primaryData[i])+Number(correctionData[i])).toFixed(2);
                    }
                    console.log(totalData);
                        console.log("создается запись");
                    var companyForMonth = new CompanyMonthCreator(totalData,"");
                    bdArray.push(companyForMonth);
                    console.log(totalData[19]);

                });
            });
            console.log(bdArray);
            payReportInputer(bdArray,usedYears,houseList,company,res);
 /*           arrayPrepairer(bdArray,usedYears,houseList,company,res)*/
        });

};

function payReportInputer(companyMonthValArray,usedYears,houseList,company,res){
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(__dirname + '/pay.xlsx')
        .then(function (payDB) {
            var worksheet = payDB.getWorksheet(1);
            companyMonthValArray.forEach(function(item, i, arr) {
                item.drainageDelta=item.drainage;
                worksheet.eachRow(function (row, rowNumber) {
                    if(item.company==row.values[1]&&item.year==row.values[2]&&item.month==row.values[3]){
                        item.payValue=row.values[4]
                        item.drainageDelta=(item.drainage-row.values[4]).toFixed(2);
                    };

                });
            })
            console.log(companyMonthValArray);
            payArrayPrepairer(companyMonthValArray,company,res)



        });

};
function payArrayPrepairer(companyMonthValArray,company,res){
    var outArray=[];
    var totalValue = new CompanyEmptyMonthCreator();
    console.log(companyMonthValArray)
    companyMonthValArray.forEach(function(item, i, arr) {
        if(item.techOmain!="0.00"){
            totalValue.adress="";
            totalValue.year="";
            totalValue.month="";
            totalValue.owner=(Number(totalValue.owner)+Number(item.owner)).toFixed(2);
            totalValue.ownersCount=(Number(totalValue.ownersCount)+Number(item.ownersCount)).toFixed(2);
            totalValue.square=(Number(totalValue.square)+Number(item.square)).toFixed(2);
            totalValue.lgotSquaree=(Number(totalValue.lgotSquaree)+Number(item.lgotSquaree)).toFixed(2);
            totalValue.techOmain=(Number(totalValue.techOmain)+Number(item.techOmain)).toFixed(2);
            totalValue.techOOther=(Number(totalValue.techOOther)+Number(item.techOOther)).toFixed(2);
            totalValue.naemSoc= (Number(totalValue.naemSoc)+Number(item.naemSoc)).toFixed(2);
            totalValue.naemKomm=(Number(totalValue.naemKomm)+Number(item.naemKomm)).toFixed(2);
            totalValue.naemBezDot= (Number(totalValue.naemBezDot)+Number(item.naemBezDot).toFixed(2));
            totalValue.heat=(Number(totalValue.heat)+Number(item.heat).toFixed(2));
            totalValue.capRepair=(Number(totalValue.capRepair)+Number(item.capRepair).toFixed(2));
            totalValue.drainage=(Number(totalValue.drainage)+Number(item.drainage)).toFixed(2);
            totalValue.waterSupply= (Number(totalValue.waterSupply)+Number(item.waterSupply)).toFixed(2);
            totalValue.eEnergy=(Number(totalValue.eEnergy)+Number(item.eEnergy)).toFixed(2);
            totalValue.compensation=(Number(totalValue.compensation)+Number(item.compensation)).toFixed(2);
            totalValue.changeSum=(Number(totalValue.changeSum)+Number(item.changeSum)).toFixed(2);
            totalValue.company=item.company;
            totalValue.recExistense=1;
            totalValue.payValue=(Number(totalValue.payValue)+Number(item.payValue)).toFixed(2);
            totalValue.drainageDelta=(Number(totalValue.drainageDelta)+Number(item.drainageDelta)).toFixed(2);


            outArray.push(item)
        }

    });
    console.log(totalValue)
    outArray.push(totalValue)
    serverResponser(outArray,company,res);
    reportCreator(outArray,company);
};

function lostIncome(company,houseList,usedYears,usedMonths,res){
    dbinputer(company,houseList,usedYears,usedMonths,res)
};

function CompanyMonthCreator(finalizedRow,houseAdress){
    this.adress=houseAdress;
    this.year=finalizedRow[2];
    this.month=finalizedRow[3];
    this.owner=finalizedRow[4];
    this.ownersCount=finalizedRow[5];
    this.square=finalizedRow[6];
    this.lgotSquaree=finalizedRow[7];
    this.techOmain=finalizedRow[8];
    this.techOOther=finalizedRow[9];
    this.naemSoc=finalizedRow[10];
    this.naemKomm=finalizedRow[11];
    this.naemBezDot=finalizedRow[12];
    this.heat=finalizedRow[13];
    this.capRepair=finalizedRow[14];
    this.drainage=finalizedRow[15];
    this.waterSupply=finalizedRow[16];
    this.eEnergy=finalizedRow[17];
    this.compensation=finalizedRow[18];
    this.changeSum=finalizedRow[19];
    this.company=finalizedRow[0];
    this.recExistense=1;
    this.payValue=0;
    this.drainageDelta=0;

}

function CompanyEmptyMonthCreator(finalizedRow,houseAdress){
    this.adress="";
    this.year="";
    this.month="";
    this.owner="";
    this.ownersCount="";
    this.square="";
    this.lgotSquaree="";
    this.techOmain="";
    this.techOOther="";
    this.naemSoc="";
    this.naemKomm="";
    this.naemBezDot="";
    this.heat="";
    this.capRepair="";
    this.drainage="";
    this.waterSupply="";
    this.eEnergy="";
    this.compensation="";
    this.changeSum="";
    this.company="";
    this.recExistense="";
    this.payValue=0;
    this.drainageDelta=0;

}

function House(finalizedRow,houseAdress){
    this.adress=houseAdress;
    this.year=finalizedRow[2];
    this.month=finalizedRow[3];
    this.owner=finalizedRow[4];
    this.ownersCount=finalizedRow[5];
    this.square=finalizedRow[6];
    this.lgotSquaree=finalizedRow[7];
    this.techOmain=finalizedRow[8];
    this.techOOther=finalizedRow[9];
    this.naemSoc=finalizedRow[10];
    this.naemKomm=finalizedRow[11];
    this.naemBezDot=finalizedRow[12];
    this.heat=finalizedRow[13];
    this.capRepair=finalizedRow[14];
    this.drainage=finalizedRow[15];
    this.waterSupply=finalizedRow[16];
    this.eEnergy=finalizedRow[17];
    this.compensation=finalizedRow[18];
    this.changeSum=finalizedRow[19];
    this.company=finalizedRow[0];
    this.recExistense=1;

}

function EmprtyHouse(houseList,company,usedYears,monthCounter){
    this.adress=houseList;
    this.year=usedYears;
    this.month=monthCounter;
    this.owner="0";
    this.ownersCount="0";
    this.square="0";
    this.lgotSquaree="0";
    this.techOmain="0";
    this.techOOther="0";
    this.naemSoc="0";
    this.naemKomm="0";
    this.naemBezDot="0";
    this.heat="0";
    this.capRepair="0";
    this.drainage="0";
    this.waterSupply="0";
    this.eEnergy="0";
    this.compensation="0";
    this.changeSum="0";
    this.company=company;
    this.recExistense=0;

}

function arrayPrepairer(bdArray,usedYears,houseList,company,res){
    var outArray = [];
    var monthCounter=1;
    var yearCounter=0;
    var stance = false;

    for (var j = 1; j < 1+12*usedYears.length; j++) {
        bdArray.forEach(function(item, i, arr) {
            if(usedYears[yearCounter]==item.year&&monthCounter==item.month&&item.company==company){
                outArray.push(bdArray[i]);
                stance=true;
            }
        });
        if(stance==false){
            var emptyhouse= new EmprtyHouse(houseList[0],company,usedYears[yearCounter],monthCounter);
            outArray.push(emptyhouse)
        };


        if(monthCounter==12){
            yearCounter++;
            monthCounter=1;
        } else{
            monthCounter++;}
        stance=false

    };
    console.log(outArray.length);
/*    serverResponser(outArray,company,res);*/
    reportCreator(outArray,company);

};

function serverResponser(outArray,company,res){
    var forResponse=JSON.stringify(outArray);
    console.log(forResponse);
    console.log("Отправлен ответ серверу");

    res.send(forResponse);
    res.end();

};
function reportCreator(outArray,company){
    console.log("Вход в функцию создания экселя осуществлен")
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(__dirname + '/exceltemplates/шаблон отчета о выпадающих доходах в адресном разрезе.xlsx')
        .then(function (reportbook) {
            var worksheet = reportbook.getWorksheet(1);
            worksheet.getCell('A6').value = 'Организация:';
            worksheet.getCell('C10').border = {
                top: {style:'thin'},
                left: {style:'thin'},
                bottom: {style:'thin'},
                right: {style:'thin'}
            };
            worksheet.getCell('F9').border = {
                top: {style:'thin'},
                left: {style:'thin'},
                bottom: {style:'thin'},
                right: {style:'thin'}
            };
            worksheet.getCell('N8').border = {
                top: {style:'thin'},
                left: {style:'thin'},
                bottom: {style:'thin'},
                right: {style:'thin'}
            };
            worksheet.getCell('A6').value ="Организация: " + company;
            outArray.forEach(function(item, i, arr) {
                worksheet.getCell('A'+(i+11)).value = i+1;
                worksheet.getCell('B'+(i+11)).value = item.adress;
                worksheet.getCell('C'+(i+11)).value = item.month+'.'+item.year;
                worksheet.getCell('D'+(i+11)).value = Number(item.techOmain);
                worksheet.getCell('E'+(i+11)).value = Number(item.techOOther);
                worksheet.getCell('F'+(i+11)).value = Number(item.naemSoc);
                worksheet.getCell('G'+(i+11)).value = Number(item.naemKomm);
                worksheet.getCell('H'+(i+11)).value = Number(item.naemBezDot);
                worksheet.getCell('I'+(i+11)).value = Number(item.heat);
                worksheet.getCell('I'+(i+11)).value = Number(item.heat);
                worksheet.getCell('J'+(i+11)).value = Number(item.capRepair);
                worksheet.getCell('K'+(i+11)).value = Number(item.drainage);
                worksheet.getCell('L'+(i+11)).value = Number(item.waterSupply);
                worksheet.getCell('M'+(i+11)).value = Number(item.eEnergy);
                worksheet.getCell('N'+(i+11)).value = Number(item.compensation);

                var row = worksheet.getRow((i+11));
                row.eachCell(function(cell, colNumber) {
                    cell.border = {
                        top: {style:'thin'},
                        left: {style:'thin'},
                        bottom: {style:'thin'},
                        right: {style:'thin'}
                    };
                });
                if(item.recExistense==0){
                    row = worksheet.getRow((i+11));
                    row.eachCell(function(cell, colNumber) {
                        cell.font = {
                            name: 'Arial',
                            color: { argb: 'FFFF0000' },
                            family: 2,
                            size: 11,
                            italic: false
                        };
                    });
                }

            });
            reportbook.xlsx.writeFile(__dirname + '/exceloutputs/отчет о выпадающих доходах в адресном разрезе.xlsx')
                .then(function() {

                    // done
                });
        })





/*function writeToExcel(usedReport,company,houseList,usedYears,usedMonths,bdArray){
    var workbook = new Excel.Workbook();
    var outArray = [];
    workbook.xlsx.readFile(__dirname + '/exceltemplates/шаблон отчета о выпадающих доходах в адресном разрезе.xlsx')
        .then(function (reportbook) {
            var emptyArr=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0];
            var startRow=0;
            var monthNumber=0;
            var worksheet = reportbook.getWorksheet(1);
            worksheet.getCell('C6').value = company;
            for (var j = 1; j < 25; j++) {

                if(j==13){
                    usedYears=2016;
                    monthNumber=0;
                }
                monthNumber++;
                var row=j+10;
                var recExist=false;
                bdArray.forEach(function(item, i, arr) {
                    if(item[1]==company&&item[3]==usedYears&&item[4]==monthNumber){
                        var selectedItem=item;
                        var a=1;
                        console.log(a)
                        console.log("Есть такая компания");
                        recExist=true;
                        outArray.push(selectedItem);}
                    worksheet.getCell('A'+row).value = j;
                    worksheet.getCell('B'+row).value = houseList;
                    worksheet.getCell('C'+row).value = monthNumber+"."+usedYears;
                    worksheet.getCell('D'+row).value = Number(selectedItem);
/!*                  worksheet.getCell('E'+row).value = Number(selectedItem[10]);



                    worksheet.getCell('A'+row).border = {
                        top: {style:'thin'},
                        left: {style:'thin'},
                        bottom: {style:'thin'},
                        right: {style:'thin'}
                    };

                });
                if(recExist==false){
                    console.log("Такой компании нет");
                    outArray.push(emptyArr);
                    worksheet.getCell('A'+row).value = j;
                    worksheet.getCell('B'+row).value = houseList;
                    worksheet.getCell('C'+row).value = monthNumber+"."+usedYears;

                }

            }

            worksheet.getRow(9).font = { name: 'Times New Roman', family: 1, size: 8, bold: false };
            worksheet.getRow(10).font = { name: 'Times New Roman', family: 1, size: 8, bold: false };
            reportbook.xlsx.writeFile(__dirname + '/exceloutputs/отчет о выпадающих доходах в адресном разрезе.xlsx')
                .then(function() {
                    // done
                });

        });
};*/





  /*  var n = fs.readFileSync('t.json',"utf-8")
    var data=JSON.parse(n);

    var useMonth=month+".xlsx";
    var workbook = new Excel.Workbook();
    workbook.creator = 'Me';
    workbook.lastModifiedBy = 'Her';
    workbook.created = new Date(1985, 8, 30);
    workbook.modified = new Date();
    var sheet = workbook.addWorksheet('Отчет', 'FFC0000');
    var worksheet = workbook.getWorksheet('Отчет');

    //заполнение шапки
    worksheet.mergeCells('A1:M1');
    worksheet.getCell('M1').value = 'Приложение 3 к договору № 0802003';
    worksheet.mergeCells('A2:M2');
    worksheet.getCell('M2').value = 'Итоговые данные отчетов';
    worksheet.mergeCells('A3:M3');
    worksheet.getCell('M3').value = 'о выпадающих доходах от предоставления льгот по оплате жилищных, коммунальных услуг,';
    worksheet.mergeCells('A4:M4');
    worksheet.getCell('M4').value = 'представленных на машинных носителях';
    worksheet.mergeCells('A5:M5');
    worksheet.getCell('M5').value = ' ';
    worksheet.mergeCells('A6:G6');
    worksheet.getCell('G6').value = 'Организация: ООО "Управляющая компания "ГрадСервис"';

    //заполнение отчета
    worksheet.mergeCells('A8:A10');
    worksheet.getCell('A10').value = '№ п/п';
    worksheet.mergeCells('B8:B10');
    worksheet.getCell('B10').value = 'Адрес дома';
    worksheet.mergeCells('C8:L8');
    worksheet.getCell('L8').value = 'Выпадающие доходы от предоставления льгот за месяц (руб.)';
    worksheet.mergeCells('C9:D9');
    worksheet.getCell('D9').value = 'Оплата жилья (техническое обслуживание)';
    worksheet.mergeCells('E9:G9');
    worksheet.getCell('G9').value = 'Плата за наём';
    worksheet.mergeCells('H9:H10');
    worksheet.getCell('H10').value = 'Отопление';
    worksheet.mergeCells('I9:I10');
    worksheet.getCell('I10').value = 'Взнос на капитальный ремонт';
    worksheet.mergeCells('J9:J10');
    worksheet.getCell('J10').value = 'Водопровод и канализации';
    worksheet.mergeCells('K9:K10');
    worksheet.getCell('K10').value = 'Горячее водоснабжение';
    worksheet.mergeCells('L9:L10');
    worksheet.getCell('L10').value = 'Электроэнергия';
    worksheet.mergeCells('M8:M10');
    worksheet.getCell('M10').value = 'Сумма компенсации за месяц (руб.)';
    worksheet.getCell('C10').value = 'на осн. площадь';
    worksheet.getCell('D10').value = 'на излишки площади';
    worksheet.getCell('E10').value = 'социальный';
    worksheet.getCell('F10').value = 'коммерческий';
    worksheet.getCell('G10').value = 'в бездотац. домах';

    console.log(data.data1.adress)
    for (var i = 1; i < 8; i++) {
        n = 10 + i;
        var number = "data" + i;
        if (data[number].adress==""){}
        else
        {
            worksheet.getCell('A' + n).value = i;
            worksheet.getCell('B' + n).value = data[number].adress;
            worksheet.getCell('C' + n).value = data[number].techOmain;
            worksheet.getCell('D' + n).value = data[number].techOOther;
            worksheet.getCell('E' + n).value = data[number].naemSoc;
            worksheet.getCell('F' + n).value = data[number].naemKomm;
            worksheet.getCell('G' + n).value = data[number].naemBezDot;
            worksheet.getCell('H' + n).value = data[number].heat;
            worksheet.getCell('I' + n).value = data[number].capRepair;
            worksheet.getCell('J' + n).value = data[number].drainage;
            worksheet.getCell('K' + n).value = data[number].waterSupply;
            worksheet.getCell('L' + n).value = data[number].eEnergy;
            worksheet.getCell('M' + n).value = data[number].compensation;
        }


    }
    workbook.xlsx.writeFile(useMonth)
        .then(function() {
            // done
        });*/

};

exports.reportSelector =reportSelector;

