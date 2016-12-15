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

class WriteData
{

	/**
	* @objPHPExcel
	*/
	public $objPHPExcel;
	
	/**
	* 规则
	*/
	public $rule0;
	public $rule1;
	public $rule2;
	public $rule3;
	public $rule4;
	public $rule5;
	public $rule6;
	public $rule7;


	/**
	* @数据
	*/
	public $excelData;

	public function __construct($rules, $excelData)
	{
		$this->rule0 = $rules[0];
		$this->rule1 = $rules[1];
		$this->rule2 = $rules[2];
		$this->rule3 = $rules[3];
		$this->rule4 = $rules[4];
		$this->rule5 = $rules[5];
		$this->rule6 = $rules[6];
		$this->rule7 = $rules[7];

		$this->excelData = $excelData;
	}

	/**
	* 写入到单元格
	*
	* @param $sheetIndex
	* @param $rules
	*
	* @return 
	*/
	private function writeToCellStyle1($sheetIndex)
	{
		$this->objPHPExcel->setActiveSheetIndex($sheetIndex);
		for( $i=0; $i < count($this->excelData); $i++ ) {
				if ( isset($this->excelData[$i]) && !empty($this->excelData[$i]) ) {
					$data = $this->excelData[$i];
					$ruleTag = 'rule'.$sheetIndex;
					foreach($this->$ruleTag as $rule) {
						$rule[0] = trim($rule[0]);
						$rule[1] = trim($rule[1]);
						if ( isset($rule[0]) && isset($rule[1]) )  {
								$fromCellName = $rule[0];
								$toCellName = $rule[1] . ($i + 2);
								$cellData = '';
								if ( isset($data[$fromCellName]) )  {
									$cellData = $data[$fromCellName];	
								}
								$this->objPHPExcel->getActiveSheet()->setCellValue($toCellName,  $cellData);
								if ( !in_array($fromCellName, ['C15', 'C16', 'C17', 'C18', 'G11']) ) {
									$this->objPHPExcel->getActiveSheet()->getStyle($toCellName)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
								}
								echo 'sheet' . $sheetIndex . '：写入字符串 “' . $cellData . '” 到 “' . $toCellName . "” 单元格，完成！\n";
						}
					}

				}			
		}

	}

	/**
	* 写入到单元格2
	*
	* @param $sheetIndex
	* @param $rules
	*
	* @return 
	*/
	private function writeToCellStyle2($sheetIndex)
	{
		$this->objPHPExcel->setActiveSheetIndex($sheetIndex);
	
		$writeLine = 2;	
		for( $i=0; $i < count($this->excelData); $i++ ) {
				if ( isset($this->excelData[$i]) && !empty($this->excelData[$i]) ) {
					$data = $this->excelData[$i];
					$ruleTag = 'rule'.$sheetIndex;
					foreach($this->$ruleTag as $key => $rules) {
						foreach($rules as $rule) {
								$rule[0] = trim($rule[0]);
								$rule[1] = trim($rule[1]);
								if ( isset($rule[0]) && isset($rule[1]) )  {
										$fromCellName = $rule[0];
										$toCellName = $rule[1] . $writeLine;
										$cellData = '';
										if ( isset($data[$fromCellName]) )  {
											$cellData = $data[$fromCellName];	
										}
										$this->objPHPExcel->getActiveSheet()->setCellValue($toCellName,  $cellData);
										if ( !in_array($fromCellName, ['C15', 'C16', 'C17', 'C18', 'G11']) ) {
											$this->objPHPExcel->getActiveSheet()->getStyle($toCellName)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
										}
										echo 'sheet' . $sheetIndex . '：写入字符串 “' . $cellData . '” 到 “' . $toCellName . "” 单元格，完成！\n";
								}
						}
						$writeLine += 1;
					}

				}			
		}
	}
	
	/**
	* 第1张表写入数据
	*
	* @param $excelData
	*
	* @return void 
	*/
	public function sheet1()
	{
		$this->writeToCellStyle1(0, $this->rule1);
	}		

	/**
	* 第2张表写入数据
	*
	* @param $excelData
	*
	* @return void 
	*/
	public function sheet2()
	{
		$this->writeToCellStyle1(1, $this->rule2);
	}
	
	/**
	* 第3张表写入数据
	*
	* @param $excelData
	*
	* @return void 
	*/
	public function sheet3()
	{
		$this->writeToCellStyle1(2, $this->rule3);
	}		

	/**
	* 第4张表写入数据
	*
	* @param $excelData
	*
	* @return void 
	*/
	public function sheet4()
	{
		$this->writeToCellStyle2(3, $this->rule4);
	}	

	/**
	* 第5张表写入数据
	*
	* @param $excelData
	*
	* @return void 
	*/
	public function sheet5()
	{
		$this->writeToCellStyle1(4, $this->rule5);
	}	

	/**
	* 第6张表写入数据
	*
	* @param $excelData
	*
	* @return void 
	*/
	public function sheet6()
	{
		$this->writeToCellStyle1(5, $this->rule6);
	}	
	
	/**
	* 第7张表写入数据
	*
	* @param $excelData
	*
	* @return void 
	*/
	public function sheet7()
	{
		$this->writeToCellStyle2(6, $this->rule7);
	}	

	/**
	* 第8张表写入数据
	*
	* @param $excelData
	*
	* @return void 
	*/
	public static function sheet8()
	{
		$this->writeToCellStyle1(7, $this->rule8);
	}	
}
