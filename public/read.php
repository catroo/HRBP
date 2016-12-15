<?php
$projectDir = dirname(dirname(__FILE__)) . '/';
include $projectDir . 'lib/phpexcel/PHPExcel.php';

function formatExcel2Array($filePath='',$sheet=0)
{
		if(empty($filePath) or !file_exists($filePath)){die('file not exists');}
		$PHPReader = new PHPExcel_Reader_Excel2007();        //建立reader对象
		if(!$PHPReader->canRead($filePath)){
				$PHPReader = new PHPExcel_Reader_Excel5();
				if(!$PHPReader->canRead($filePath)){
						echo 'no Excel';
						return ;
				}
		}
		$PHPExcel = $PHPReader->load($filePath);        //建立excel对象
		$currentSheet = $PHPExcel->getSheet($sheet);        //**读取excel文件中的指定工作表*/
		$allColumn = $currentSheet->getHighestColumn();        //**取得最大的列号*/
		echo 'allColumn' . $allColumn . "\n";
		$allRow = $currentSheet->getHighestRow();        //**取得一共有多少行*/
		$data = array();
		for($rowIndex=1;$rowIndex<=$allRow;$rowIndex++){        //循环读取每个单元格的内容。注意行从1开始，列从A开始
				for($colIndex='A';$colIndex<=$allColumn;$colIndex++){
						$addr = $colIndex.$rowIndex;
						$cell = $currentSheet->getCell($addr)->getValue();
						echo $colIndex . '-->' . $currentSheet->getColumnDimension($colIndex)->getWidth() . "\n";
						if($cell instanceof PHPExcel_RichText){ //富文本转换字符串
								$cell = $cell->__toString();
						}
						$data[$rowIndex][$colIndex] = $cell;
				}
		}
		return $data;
}

function getCurrectSheetCellWidth($filePath='',$sheet=0)
{
		if(empty($filePath) or !file_exists($filePath)){die('file not exists');}
		$PHPReader = new PHPExcel_Reader_Excel2007();        //建立reader对象
		if(!$PHPReader->canRead($filePath)){
				$PHPReader = new PHPExcel_Reader_Excel5();
				if(!$PHPReader->canRead($filePath)){
						echo 'no Excel';
						return ;
				}
		}
		$PHPExcel = $PHPReader->load($filePath);        //建立excel对象
		$currentSheet = $PHPExcel->getSheet($sheet);        //**读取excel文件中的指定工作表*/
		$allColumn = $currentSheet->getHighestColumn();        //**取得最大的列号*/
		echo 'allColumn' . $allColumn . "\n";
		$allRow = $currentSheet->getHighestRow();        //**取得一共有多少行*/
		$data = array();
		for($colIndex='A';$colIndex<=$allColumn;$colIndex++){
				echo $colIndex . '-->' . $currentSheet->getColumnDimension($colIndex)->getWidth() . "\n";
		}
}



$filePath = $projectDir . 'public/1.xlsx';
$data = getCurrectSheetCellWidth($filePath, 0);
exit;
