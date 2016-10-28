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
