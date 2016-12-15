<?php
/* Copyright (C) 
* 2016 - dinghui@plu.cn
* This program is free software; you can redistribute it and/or
* modify it under the terms of the GNU General Public License
* as published by the Free Software Foundation; either version 2
* of the License, or (at your option) any later version.
* 
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
* GNU General Public License for more details.
* 
* You should have received a copy of the GNU General Public License
* along with this program; if not, write to the Free Software
* Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
* 
*/

$projectDir = dirname(dirname(__FILE__)) . '/';
include $projectDir . 'lib/phpexcel/PHPExcel.php';
include $projectDir . 'lib/phpexcel/PHPExcel/Writer/Excel2007.php';
include $projectDir . 'conf/config.php';
include $projectDir . 'lib/utils.php';
include $projectDir . 'lib/readData.php';
include $projectDir . 'lib/writeData.php';

$resourceDir = $projectDir . '/resource';

// 列出所有员工信息表
$files = ReadData::loadExcelFiles($resourceDir);
echo '共找到' . count($files) . '个员工登记表！' . "\n";

// 把员工信息表存储到数组
$excelData = [];
foreach($files as $file) {
	$res = ReadData::loadExcelData($file);
	if ( $res ) {
		$excelData[] = $res; 
	}
}

$objPHPExcel = new PHPExcel();
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);

// 初始化工作表
for($i=0; $i < count($sheets); $i++) {
		
		echo '正在处理工作表：' . $sheets[$i]['name'] . "\n";
		
		$objPHPExcel->setActiveSheetIndex($i);
		$objPHPExcel->getActiveSheet()->setTitle($sheets[$i]['name']);
		$objPHPExcel->createSheet();

		foreach($sheets[$i]['title'] as $key => $val) {
			
			// 格式化cellName
			$ascii = ord('A');
			$ascii = $ascii + $key;
			$alphabet = Utils::formatCellName($ascii);
			$addr = $alphabet . '1';
		
		
			// 处理CellName数据
			if ( is_string($val) ) {
				$cellNameValue = $val;	
				$cellNameWidth = null;	
			} else if ( is_array($val) )  {
				$cellNameValue = $val[0];	
				$cellNameWidth = (int)$val[1]/6;	
			}
			
			// 设置title
			$objPHPExcel->getActiveSheet()->setCellValue($addr,  $cellNameValue);
			$objPHPExcel->getActiveSheet()->getStyle($addr)->getFont()->setBold(true);
			$objPHPExcel->getActiveSheet()->getStyle($addr)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
			$objPHPExcel->getActiveSheet()->getStyle($addr)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
			$objPHPExcel->getActiveSheet()->getStyle($addr)->getFont()->setSize(9);

			// 设置宽度
			if ( $cellNameWidth != null ) {
				$objPHPExcel->getActiveSheet()->getColumnDimension($alphabet)->setWidth($cellNameWidth);
			} else {
				$objPHPExcel->getActiveSheet()->getColumnDimension($alphabet)->setAutoSize(true);
			}
		}

		// 载入数据规则
		$rules[$i] = $sheets[$i]['rule'];
}

// 数据写入到excel
$writeData = new WriteData($rules, $excelData);
$writeData->objPHPExcel = $objPHPExcel;
$writeData->sheet1();
$writeData->sheet2();
$writeData->sheet3();
$writeData->sheet4();
$writeData->sheet5();
$writeData->sheet6();
$writeData->sheet7();

$objWriter->save($projectDir . "/completed/done.xlsx");
