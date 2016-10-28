<?php
/**
 * Created by PhpStorm.
 * User: Abo
 * Date: 2016/1/6
 * Email:772764794@qq.com
 */
include "./PHPExcel/Classes/PHPExcel.php"; // dirname(__FILE__).$path;//--E:\Deve\xampp\htdocs\Tools/PHPExcel/Classes/PHPExcel.php
include "./PHPExcel/Classes/PHPExcel/IOFactory.php";   //引入读取excel的类文件

class cls_excel{
    private $charIndex; //列名大写字母下标

    public function __construct(){
        $this->charIndex=range('A','Z');
    }


    /**
     * 导出Excel(需创建单表信息)
     * @param string $fileName 文件名
     * @param string $sheetName 单表名
     * @param array $cellNames 列名s （一维数组）
     * @param array $data 导出数据（二维数组，英文下标）array(0=>array('id'=>1,'cate_id'=>12,'name'=>'abc'))
     * @param int $sheetIndex 单表下标，默认0
     */
    public function CreateSheetInfo_Export($fileName,$sheetName,$cellNames,$cellWidth,$data,$sheetIndex=0){
        $this->charIndex=range('A','Z');
        $cell_names2sheet=array();
        foreach($cellNames as $key=>$val){
            $cell_names2sheet[$this->charIndex[$key].'1']=$val;
        }

        $sheet=[
            'fileName'=>$fileName,
            'sheetIndex'=>$sheetIndex,
            'title'=>$sheetName,
            'cellName'=>$cell_names2sheet,   //Excel列名
            'cellWidht'=>$cellWidth,//列宽
            'cellIndex'=>array_keys($data[0]),  //excel导入查询数据用 引导下标
            'data'=>$data,
        ];
        return $this->ExportExcel($sheet);
    }

/**
     * 输出二维数组
     * @param $sheet 二维数组
     * @throws PHPExcel_Reader_Exception
     */
    private function ExportExcel($sheet){
        try {
            $objPHPExcel = new PHPExcel();  //实例化PHPExcel类， 等同于在桌面上新建一个excel
            $objPHPExcel->setActiveSheetIndex($sheet['sheetIndex']);//把新创建的sheet设定为当前活动sheet
            $objSheet = $objPHPExcel->getActiveSheet();//获取当前活动sheet
            $objSheet->setTitle($sheet['title']);//给当前活动sheet起个名称

            $i2cell=0;
            foreach ($sheet['cellName'] as $k => $v) {
                $objPHPExcel->getActiveSheet()->getColumnDimension($this->charIndex[$i2cell])->setWidth($sheet['cellWidht'][$i2cell]);//设置列宽
                $objSheet = $objSheet->setCellValue($k, $v);//设表名
                $i2cell++;
            }

            $j = 2;   //行数
            foreach ($sheet['data'] as $val) {
                $i = 0; //列数&对填充的查询字段
                while ($i < sizeof($sheet['cellName'])) {
                    $objSheet = $objSheet->setCellValue($this->charIndex[$i] . $j, $val[$sheet['cellIndex'][$i]]);
                    $i++;
                }
                $j++;
            }

            ob_end_clean();//清除缓冲区,避免乱码
            $objWriter=PHPExcel_IOFactory::createWriter($objPHPExcel,'Excel5');//生成excel文件
            //$objWriter->save($dir."/export_1.xls");//保存文件
            $this->browser_export('Excel5',"{$sheet['fileName']}.xls");//输出到浏览器
            $objWriter->save("php://output");
            return true;
        } catch (Exception $e) {
            echo json_encode(['status'=>false,'msg'=>'系统异常：' . $e->getMessage()]);
        }
    }

    /**
     * 输出到浏览器
     * @param $type excel类型
     * @param $filename 输出文件名
     */
    private function browser_export($type,$filename){
        if($type=="Excel5"){
            header("Content-type:text/html;charset=utf-8");
            header('Content-Type: application/vnd.ms-excel');//告诉浏览器将要输出excel03文件
        }else{
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');//告诉浏览器数据excel07文件
        }
        header('Content-Disposition: attachment;filename="'.$filename.'"');//告诉浏览器将输出文件的名称
        header('Cache-Control: max-age=0');//禁止缓存
    }

    /**
     * 导入Excel表
     * @param string $filepath 文件所在路径
     * @param string $sheetName 表名
     * @param array $assocIndex 字段下标值(一维数组) array('id','cat_id','name','sale','details','add_time');
     * @return $arr 返回导入Excel数组
     */
    public function inputExcel($filepath,$sheetName,$assocIndex){
        header("Content-Type:text/html;charset=utf-8");
        $filename=$filepath;//找到当前脚本所在路径
        $sheetName=array("{$sheetName}");
        try {
            $fileType = PHPExcel_IOFactory::identify($filename);//自动获取文件的类型提供给phpexcel用
            $objReader = PHPExcel_IOFactory::createReader($fileType);//获取文件读取操作对象
            $objReader->setLoadSheetsOnly($sheetName);//只加载指定的sheet
            $objPHPExcel = $objReader->load($filename);//加载文件

            $sheet = $objPHPExcel->getActiveSheet();  //取一张表内容
            foreach ($sheet->getRowIterator() as $hang => $row) {//逐行处理
                if ($row->getRowIndex() < 2) {continue;}
                $i = 0;
                foreach ($row->getCellIterator() as $lie => $cell) {//逐列读取
                    $arr[$hang][$assocIndex[$i]] = $cell->getValue();//获取单元格数据
                    $i++;
                }
            }
        }catch (Exception $e){
            echo json_encode(['status'=>'false','msg'=>$e->getMessage()]);
        }
        return $arr;
    }
}

//导出DEMO

header("Content-type: text/html; charset=utf-8");
$et=new cls_excel();
//$et->ExportExcel($sheet);

$sheet_name="球场佣金报表(20160701_20161028)";
$cellname = ['订单号','球场名称','打球日期', '预定人', '打球人数','套餐单价','订单总额', "佣金", '付款方式', '供应商','所属区域','新客记录'];
$cellWidth = [20,50,10, 10, 10, 10, 20,10, 10,30,10,10];
$goodsList='[{"order_sn":" 201610161056524704","golf_course_name":"\u5e7f\u5dde\u9e93\u6e56\u9ad8\u5c14\u592b\u7403\u4e61\u6751\u4ff1\u4e50\u90e8","bdate":"2016-10-17","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"990.00","total_amount":"990.00","commission":"40.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609231206391546","golf_course_name":"\u4e1c\u839e\u51e4\u51f0\u5c71\u9ad8\u5c14\u592b\u4ff1\u4e50\u90e8","bdate":"2016-09-23","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"750.00","total_amount":"750.00","commission":"30.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u662f"},{"order_sn":" 201609231449335323","golf_course_name":"\u5e7f\u5dde\u9e93\u6e56\u9ad8\u5c14\u592b\u7403\u4e61\u6751\u4ff1\u4e50\u90e8","bdate":"2016-09-23","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"890.00","total_amount":"890.00","commission":"40.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609231452485005","golf_course_name":"\u4e1c\u839e\u51e4\u51f0\u5c71\u9ad8\u5c14\u592b\u4ff1\u4e50\u90e8","bdate":"2016-09-23","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"750.00","total_amount":"750.00","commission":"30.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609231454016896","golf_course_name":"\u5e7f\u5dde\u5357\u6c99\u9ad8\u5c14\u592b-\u6e56\u573a(AB\u573a)","bdate":"2016-09-23","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"705.00","total_amount":"705.00","commission":"30.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u5e7f\u5dde\u8c6a\u5bcc","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609191747529553","golf_course_name":"\u5e7f\u5dde\u9e93\u6e56\u9ad8\u5c14\u592b\u7403\u4e61\u6751\u4ff1\u4e50\u90e8","bdate":"2016-09-20","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"760.00","total_amount":"760.00","commission":"40.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u6df1\u5733\u94c1\u9a6c","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u662f"},{"order_sn":" 201609191907448623","golf_course_name":"\u4e2d\u5c71\u6e29\u6cc9\u9ad8\u5c14\u592b(\u65e7\u573a)","bdate":"2016-09-20","player_name":"\u6d4b\u8bd5","number_of_players":"3","single_price":"800.00","total_amount":"2400.00","commission":"0.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"\u4e2d\u5c71\u6e29\u6cc9\u9ad8\u5c14\u592b(\u65e7\u573a)","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609201429246786","golf_course_name":"\u6df1\u5733\u805a\u8c6a\u4f1a\u65b0\u573aHEF","bdate":"2016-09-20","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"590.00","total_amount":"590.00","commission":"0.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"},{"order_sn":" 201609071551191591","golf_course_name":"\u4e1c\u839e\u51e4\u51f0\u5c71\u9ad8\u5c14\u592b\u4ff1\u4e50\u90e8","bdate":"2016-09-08","player_name":"\u5c0f\u670b\u53cb","number_of_players":"1","single_price":"750.00","total_amount":"750.00","commission":"30.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u662f"},{"order_sn":" 201607261139327405","golf_course_name":"\u5ee3\u5dde\u9e93\u6e56\u9ad8\u723e\u592b\u7403\u9109\u6751\u4ff1\u6a02\u90e8","bdate":"2016-07-27","player_name":"\u6d4b\u8bd5","number_of_players":"1","single_price":"760.00","total_amount":"760.00","commission":"40.00","payment_type":"\u7403\u573a\u73b0\u4ed8","agent_name":"","region":"\u534e\u5357\u5730\u533a","is_first_order":"\u5426"}]';
$et->CreateSheetInfo_Export($sheet_name,$sheet_name,$cellname,$cellWidth,json_decode($goodsList,true));


/*
//导入DEMO
$et=new cls_excel();


$filepath=dirname(__FILE__).'/20160701_20161028.xls';
$sheet_name="球场佣金报表(20160701_20161028)";
$assocIndex=['订单号','球场名称','打球日期', '预定人', '打球人数','套餐单价','订单总额', "佣金", '付款方式', '供应商','所属区域','新客记录'];
$arr=$et->inputExcel($filepath,$sheet_name,$assocIndex);
echo '<pre>';
var_dump($arr);
*/
