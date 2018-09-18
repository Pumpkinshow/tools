#!/usr/bin/env node
var program = require('commander');
var bx = require('../xlsx.js');
var Path = require("path");

program
.version('0.0.1')
.usage('test')
.option('-o, --origin [value]', '设置原文件名字')
.option('-d, --dist [value]', '设置产出地址','./dist.xlsx')
// .option('-m, --max <n>', '最大连接数', parseInt)
// .option('-s, --seed <n>', '出始种子', parseFloat)

program
.command('pack-bx <command> [<args>] [<value>]')
.description('一个好房报销工具')
.action(function(name){
console.log('Deploying "%s"', name);
});

program.parse(process.argv);

console.log('源文件地址：', Path.resolve(process.cwd(), program.origin));
console.log('产出地址：', Path.resolve(process.cwd(), program.dist));

if(!program.origin){
    throw new Error("报销源文件必须写，详情请查看帮助，pack-bx -help");
    return;
}
var reg = /.+(\.xlsx)$/;
var dist = program.dist;
if(dist && !reg.test(dist)){
    dist += '.xlsx';
}
// console.log(__dirname); //code地址
// console.log(process.cwd()); //命令行pwd地址
// console.log(process.execPath);//node.exe 地址
bx(Path.resolve(process.cwd(),program.origin),Path.resolve(process.cwd(),dist));


