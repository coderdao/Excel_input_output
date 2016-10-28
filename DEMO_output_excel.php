<?php

//导出DEMO

header("Content-type: text/html; charset=utf-8");
$et=new cls_excel();

$sheet_name="球场佣金报表(20160701_20161028)";
$cellname = ['订单号','球场名称','打球日期', '预定人', '打球人数','套餐单价','订单总额', "佣金", '付款方式', '供应商','所属区域','新客记录'];
$cellWidth = [20,50,10, 10, 10, 10, 20,10, 10,30,10,10];
$goodsList='[{"order_sn":" 201610161056524704","golf_course_name":"\u5e7f\u5dde\u9e93\u6e56\u9ad8\u5c14\u592b\u7403\u4e61\u6751\u4ff1\u4e50\u90e8","bdate":"2016-10-17","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"990.00","total_amount":"990.00","commission":"40.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609231206391546","golf_course_name":"\u4e1c\u839e\u51e4\u51f0\u5c71\u9ad8\u5c14\u592b\u4ff1\u4e50\u90e8","bdate":"2016-09-23","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"750.00","total_amount":"750.00","commission":"30.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u662f"},{"order_sn":" 201609231449335323","golf_course_name":"\u5e7f\u5dde\u9e93\u6e56\u9ad8\u5c14\u592b\u7403\u4e61\u6751\u4ff1\u4e50\u90e8","bdate":"2016-09-23","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"890.00","total_amount":"890.00","commission":"40.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609231452485005","golf_course_name":"\u4e1c\u839e\u51e4\u51f0\u5c71\u9ad8\u5c14\u592b\u4ff1\u4e50\u90e8","bdate":"2016-09-23","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"750.00","total_amount":"750.00","commission":"30.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609231454016896","golf_course_name":"\u5e7f\u5dde\u5357\u6c99\u9ad8\u5c14\u592b-\u6e56\u573a(AB\u573a)","bdate":"2016-09-23","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"705.00","total_amount":"705.00","commission":"30.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609191747529553","golf_course_name":"\u5e7f\u5dde\u9e93\u6e56\u9ad8\u5c14\u592b\u7403\u4e61\u6751\u4ff1\u4e50\u90e8","bdate":"2016-09-20","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"760.00","total_amount":"760.00","commission":"40.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u6df1\u5733\u94c1\u9a6c","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u662f"},{"order_sn":" 201609191907448623","golf_course_name":"\u4e2d\u5c71\u6e29\u6cc9\u9ad8\u5c14\u592b(\u65e7\u573a)","bdate":"2016-09-20","player_name":"\u6d4b\u8bd5","number_of_players":"3","single_price":"800.00","total_amount":"2400.00","commission":"0.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u4e2d\u5c71\u6e29\u6cc9\u9ad8\u5c14\u592b(\u65e7\u573a)","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609201429246786","golf_course_name":"\u6df1\u5733\u805a\u8c6a\u4f1a\u65b0\u573aHEF","bdate":"2016-09-20","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"590.00","total_amount":"590.00","commission":"0.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609071551191591","golf_course_name":"\u4e1c\u839e\u51e4\u51f0\u5c71\u9ad8\u5c14\u592b\u4ff1\u4e50\u90e8","bdate":"2016-09-08","player_name":"\u5c0f\u670b\u53cb","number_of_players":"1","single_price":"750.00","total_amount":"750.00","commission":"30.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u662f"},{"order_sn":" 201607261139327405","golf_course_name":"\u5ee3\u5dde\u9e93\u6e56\u9ad8\u723e\u592b\u7403\u9109\u6751\u4ff1\u6a02\u90e8","bdate":"2016-07-27","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"760.00","total_amount":"760.00","commission":"40.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"}]';

$et->CreateSheetInfo_Export($sheet_name,$sheet_name,$cellname,$cellWidth,json_decode($goodsList,true));
