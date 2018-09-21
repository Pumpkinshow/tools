// https://www.npmjs.com/package/xlsx
// https://www.npmjs.com/package/xlsx-style

var XLS = require('xlsx');
var fs = require('fs');
var XLSX = require('xlsx-style');
var PATH = require("path");

 module.exports = function(o,d){
    var origin = o;
    var dist = d;
    //临时文件存放的目录
    var tempUrl = PATH.resolve(__dirname , './temp.xlsx');
    // xlsx读取xls文件
    var workbook_xls = XLS.readFile(origin,{cellStyles:true,cellDates:false});
    // 生成temp.xlsx的原因是，xlsx-style插件不能识别.xls的文件，需生成一个临时temp.xlsx
    XLS.writeFile(workbook_xls, tempUrl, {bookType:'xlsx',bookSST: false, type: 'binary'});
    var workbook = XLSX.readFile(tempUrl,{cellStyles:true,cellDates:false});
    // {
    //     SheetNames: ['sheet1', 'sheet2'],
    //     Sheets: {
    //         // worksheet
    //         'sheet1': {
    //             // cell
    //             'A1': { ... },
    //             // cell
    //             'A2': { ... },
    //             ...
    //         },
    //         // worksheet
    //         'sheet2': {
    //             // cell
    //             'A1': { ... },
    //             // cell
    //             'A2': { ... },
    //             ...
    //         }
    //     }
    // }

    // 加班容器
    var jb = {};
    var jb_keys = [];
    var Sheets = workbook.Sheets;
    var newSheets = {};
    var reg = /="(.+)"/;
    var name = "";
    for(var key in Sheets){
        for(var k in Sheets[key]){
            //="xxxx" => xxxx,去掉两边的双引号
            if(reg.test(Sheets[key][k]['v'])){
                Sheets[key][k]['v'] = Sheets[key][k]['v'].replace(reg,'$1');
            }
            // 获取加班日期，存储为keys
            if(Sheets[key][k]['s'] && Sheets[key][k]['s']['numFmt'] =='m/d/yy'){
                var date = XLSX.SSF.parse_date_code(Sheets[key][k]['v']);
                if([1,2,3,4,5].indexOf(date.q) ==-1){
                    jb[date.y + '-' + date.m + '-' + date.d] = 1;
                }else{
                    if(date.H>20 || (date.H==20 && date.M>=30)){
                        jb[date.y + '-' + date.m + '-' + date.d] = 1;
                    }
                }
            }
        }
    }
    console.log(JSON.stringify(workbook.Sheets).substr(0,2000));
    // 获取keys
    jb_keys = Object.keys(jb);
    // 修改workbook
    for(var m in Sheets){
        var i = 1, 
            j = 0, 
            sheetArr = [],
            sheet={},
            item,
            i_i=1,
            sameDay,
            repay=0,
            repayTotal=0;
        // 合并的数组
        var merges = [];
        for(var n in Sheets[m]){
            if(n == '!ref') continue;
            item = Sheets[m][n];
            // console.log(item);
            j>=8 ? j=0 : "";
            j == 0 ? sheetArr = [item] : sheetArr.push(item);
            if(j==7){
                var date = XLSX.SSF.parse_date_code(sheetArr[3]['v']),firstday;
                // console.log(date);
                (date.y != '1900') && !sameDay && (jb_keys.indexOf(date.y + '-' + date.m + '-' + date.d) != -1) &&(sameDay = sheetArr[3]['v']);
                if(sameDay && (firstday = XLSX.SSF.parse_date_code(sameDay)) && jb_keys.indexOf(date.y + '-' + date.m + '-' + date.d) != -1 && (firstday.m == date.m && firstday.d != date.d)){
                    merges.push({s: {c: 5, r:i_i }, e: {c:5, r:i-2}});
                    merges.push({s: {c: 6, r:i_i }, e: {c:6, r:i-2}});
                    merges.push({s: {c: 7, r:i_i }, e: {c:7, r:i-2}});
                    repayTotal += repay;
                    sameDay = sheetArr[3]['v'];
                    i_i = i-1;
                }

                if(date.y == "1900" || jb_keys.indexOf(date.y + '-' + date.m + '-' + date.d) != -1){
                    sheet['A' + i] = sheetArr[0];
                    sheet['B' + i] = sheetArr[1];
                    sheet['C' + i] = sheetArr[2];
                    sheet['D' + i] = sheetArr[3];
                    sheet['D' + i]['z'] = "yyyy-mm-dd hh:mm";
                    sheet['E' + i] = sheetArr[4];
                    repay = 25;
                    if([1,2,3,4,5].indexOf(date.q) ==-1){
                        repay = 50;
                    }
                    name = sheetArr[1].v;
                    if(date.y == '1900'){
                        sheet['F' + i] = {"t":"s","v":"加班误餐费"};
                        sheet['G' + i] = {"t":"s","v":"交通费"};
                        sheet['H' + i] = {"t":"s","v":"合计"};
                    }else{
                        sheet['F' + i] = {"t":"n","v":repay};
                        sheet['G' + i] = {"t":"n","v":0};
                        sheet['H' + i] = {"t":"n","v":repay};
                    }
                    i++;
                }
            }
            j++;
        }
        repayTotal += repay;
        sheet['A' + i] = {"t":"s","v":'合计','s':{alignment:{vertical:'center',horizontal:'center'}}};
        sheet['B' + i] = {"t":"s","v":'','s':{alignment:{vertical:'center',horizontal:'center'}}};
        sheet['C' + i] = {"t":"s","v":'','s':{alignment:{vertical:'center',horizontal:'center'}}};
        sheet['D' + i] = {"t":"s","v":'','s':{alignment:{vertical:'center',horizontal:'center'}}};
        sheet['E' + i] = {"t":"s","v":'','s':{alignment:{vertical:'center',horizontal:'center'}}};
        sheet['F' + i] = {"t":"n","v":repayTotal,'s':{alignment:{vertical:'center',horizontal:'center'}}};
        sheet['G' + i] = {"t":"n","v":0,'s':{alignment:{vertical:'center',horizontal:'center'}}};
        sheet['H' + i] = {"t":"n","v":repayTotal,'s':{alignment:{vertical:'center',horizontal:'center'}}};

        for(var ks in sheet){
            sheet[ks]['s'] = {
                alignment:{vertical:'center',horizontal:'center'},
                border:{
                    top: { style: "thin", color: { auto: 1} },
                    bottom: { style: "thin", color: { auto: 1} },
                    left: { style: "thin", color: { auto: 1} },
                    right: { style: "thin", color: { auto: 1} }
                }
            };
        }

        sheet['F' + (i+2)] = {"t":"s","v":"报销人：" + name ,'s':{alignment:{vertical:'center'}}};
        sheet['F' + (i+3)] = {"t":"s","v":"报销金额：" + repayTotal + "元",'s':{alignment:{vertical:'center'}}};
        sheet['F' + (i+4)] = {"t":"s","v":"餐饮发票：0元",'s':{alignment:{vertical:'left'}}};
        sheet['F' + (i+5)] = {"t":"s","v":"交通费发票：0元",'s':{alignment:{vertical:'left'}}};
        merges.push({s: {c: 5, r:i_i }, e: {c:5, r:i-2}});
        merges.push({s: {c: 6, r:i_i }, e: {c:6, r:i-2}});
        merges.push({s: {c: 7, r:i_i }, e: {c:7, r:i-2}});
        // merges.unshift({s: {c: 2, r:1 }, e: {c:2, r:i-2}});
        merges.unshift({s: {c: 1, r:1 }, e: {c:1, r:i-2}});
        merges.unshift({s: {c: 0, r:1 }, e: {c:0, r:i-2}});

        
        newSheets[m] = sheet;
        // 展示区域需要预先定义
        newSheets[m]['!ref'] = 'A1:H' + (i+5);
        //合并单元格数组
        newSheets[m]['!merges'] = merges;
        newSheets[m]['!cols'] = [{wch:15},{wch:10},{wch:40},{wch:20},{wch:5},{wch:10},{wch:10}];
    }
    // 一个SheetName数组
    var sheetNames = workbook.SheetNames;

    // 新的workbook
    var newWorkbook = {
        SheetNames:sheetNames,
        Sheets:newSheets
    };

    // console.log(JSON.stringify(newWorkbook).substr(0,2000));
    // console.log(workbook.SheetNames);
    // console.log(workbook.Sheets[sheetNames[0]]['!ref']) //有效范围

    //删除temp.xlsx 文件
    fs.unlink(tempUrl);
    XLSX.writeFile(newWorkbook, dist,{cellDates:true});
}