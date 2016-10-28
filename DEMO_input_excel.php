<?php
//导入DEMO
include './cls_excel.php';
$et=new cls_excel();

$filepath=dirname(__FILE__).'/20160701_20161028.xls';
$sheet_name="球场佣金报表(20160701_20161028)";
$assocIndex=['订单号','球场名称','打球日期', '预定人', '打球人数','套餐单价','订单总额', "佣金", '付款方式', '供应商','所属区域','新客记录'];

$arr=$et->inputExcel($filepath,$sheet_name,$assocIndex);

echo '<pre>';
var_dump($arr);
