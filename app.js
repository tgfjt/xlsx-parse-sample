'use strict';

var fs = require('fs');

var XLSX = require('xlsx');
var Handlebars = require('handlebars');
var _ = require('lodash');
var moment = require('moment');

var xlsxPath = 'resources/sample.xlsx';
var hbsPath = 'view/index.hbs';
var copyrightPath = 'view/copyright.hbs';
var htmlPath = 'public/html/index.html';
var template = Handlebars.compile(fs.readFileSync(hbsPath, 'utf8'));
Handlebars.registerPartial('copyright', fs.readFileSync(copyrightPath, 'utf8'));

var file = XLSX.readFile(xlsxPath);
var target = file.Sheets.JS;

var lists = _.chain(Object.keys(target).splice(15))
  .groupBy(function(e, i) {
    return Math.floor(i / 6);
  })
  .toArray()
  .value();

var rows = _.map(lists, function(a) {
  return {
    number: target[a[0]].v,
    id: target[a[1]].v,
    name: target[a[2]].v,
    category: target[a[3]].v,
    start: moment(target[a[4]].w, 'MM/DD/YY', 'ja').format('LL'),
    end: moment(target[a[5]].w, 'MM/DD/YY', 'ja').format('LL')
  };
});

var result = template({
  title: 'とあるJSのスケジュール',
  update: moment().format('YYYY-MM-DD'),
  row: rows
});


var ws = fs.createWriteStream(htmlPath)
  .on('error', function(err) {
    console.log(err);
  })
  .on('end', function() {
    console.log(htmlPath, '保存が完了しました。');
  });

ws.write(result);
ws.end();
