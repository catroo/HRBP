<?php
/* Copyright (C) 
* 2016 - dinghui
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

class ReadData {

		/**
		* 获取所有待整理的excel
		*
		* @param $dir
		*
		* @return array
		*/
		public static function loadExcelFiles($dir)
		{
				$files = [];
				if(is_file($dir)) {
						return $dir;
				}
				$handle = opendir($dir);
				if( $handle ) {
						while(false !== ($file = readdir($handle))) {
								if ($file != '.' && $file != '..' && $file != '.DS_Store' ) {
										$ext = strtolower(substr(strrchr($file, '.'), 1));
										$filename = $dir . '/' . $file;
										
										if( is_dir($filename) ) {
											$files = array_merge($files, self::loadExcelFiles($filename));
										} else {
											if ( $ext != 'xlsx' && $ext != 'xls' ) {
												Utils::errorLog('文件格式不正确跳过：' . $filename);
												continue;
											}
											$files[] = $filename;
										}

								}
						}
						closedir($handle);
				}
				return $files;
		}

		/**
        * 加载员工信息表excel信息
		*
		* @param $filePath
		* @param $sheet
		*
		* @return 
		*/
		public static function loadExcelData($filePath='', $sheet=0)
        {
				if( empty($filePath) || !file_exists($filePath) ) {
					return;
				}
				
				// 建立reader对象
				$phpReader = new PHPExcel_Reader_Excel2007();
				if( !$phpReader->canRead($filePath) ) {
						$phpReader = new PHPExcel_Reader_Excel5();
						if( !$phpReader->canRead($filePath) ) {
								Utils::errorLog('Excel格式不正确:' . $filePath);
								return;
						}
				}

        		// 建立excel对象
				$phpExcel = $phpReader->load($filePath);
				echo '正在读取：' . $filePath . "\n";

				// 读取excel文件中的指定工作表
				$currentSheet = $phpExcel->getSheet($sheet);

				// 取得最大的列号
				$allColumn = $currentSheet->getHighestColumn();

				// 取得一共有多少行
				$allRow = $currentSheet->getHighestRow();
				$data = array();

				// 循环读取每个单元格的内容。注意行从1开始，列从A开始
				for($rowIndex=2; $rowIndex<=$allRow; $rowIndex++) {
					for($colIndex='A'; $colIndex<=$allColumn; $colIndex++) {
						$addr = $colIndex.$rowIndex;
						$cellObj = $currentSheet->getCell($addr);
						$cellValue = $cellObj->getValue();
						// 富文本转换字符串
						if( $cellValue instanceof PHPExcel_RichText ) {
								$cellValue = $cellValue->__toString();
						}
						$data[$addr] = $cellValue;
					}
				}

				// 校验关键信息
				$verified = self::verifyData($data);
				if ( $verified == false ) {
					Utils::errorLog('文档非标准模板格式或者缺少关键信息：' . $filePath);
					return false;
				}

				return self::formatExcelData($data);
		}


		/**
		* 校验数据格式
		*
		* @param $data
		*
		* @return 
		*/
		public static function verifyData($data)
		{
			$verified = true;

			// 姓名缺失
			if ( empty($data['C2']) ) {
				$verified = false;
			}
		
			// 身份证缺失
			if ( empty($data['C6']) ) {
				$verified = false;
			}

			//	个人手机号
			if ( empty($data['C11']) ) {
				$verified = false;
			}

			// 户口地址
			if ( empty($data['C15']) ) {
				$verified = false;		
			}
			return $verified;
		}
		
		/**
		* 格式化excel数据
		*
		* @param $data
		*
		* @return array 
		*/
		public static function formatExcelData($data)
		{
			foreach($data as $key => $val) {
				
				if ( empty($val) ) {
					unset($data[$key]);		
				}

				$data[$key] = trim($val);

				// 格式化这几个日期格式的数据
				if ( in_array($key, ['E3', 'I8', 'D20', 'D21', 'I42']) ) {
					$data[$key] = trim($data[$key]);
					if ( is_string($val) ) {
						$data[$key] = date('Y-m-d', strtotime($val));		
					}
					else if ( is_numeric($val) ) {
						if ( strlen($val) == 8 ) {
							$data[$key] = $val{4} . '-' . $val{2} . '-' . $val{2};	
						} else {
							$timestamp = intval(($val - 25569) * 3600 * 24);
							$data[$key] = gmdate('Y-m-d', $timestamp);	
						}
					}
				}	

				// 格式化国家
				if ( $key == 'E4' && in_array($val, ['中华人民共和国', '中国国籍', '汉']) ) {
					$data[$key] = '中国';	
				}

				// 格式化地址
				if ( in_array($key, ['C15', 'C16', 'C17', 'C18', 'G6']) ) {
					$data[$key] = preg_replace('/_|\s/', '', $data[$key]);	

					if ( strpos($data[$key], '上海') !== false || strpos($data[$key], '北京') !== false ) {
						$data[$key] = str_replace('省', '', $data[$key]);		
					}
					if ( (strpos($data[$key], '上海') !== false || strpos($data[$key], '北京') !== false ) && strpos($data[$key], '区') > 0 ) {
						$data[$key] = str_replace('区/县', '', $data[$key]);		
					}

					if ( strpos($data[$key], '北京北京') !== false ) {
						$data[$key] = str_replace('北京北京', '北京', $data[$key]);		
					} 

					if ( strpos($data[$key], '上海上海') !== false ) {
						$data[$key] = str_replace('上海上海', '上海', $data[$key]);		
					} 

					if ( strpos($data[$key], '省省') !== false ) {
						$data[$key] = str_replace('省省', '省', $data[$key]);		
					} 
					
					$data[$key] = str_replace('市区/县', '市', $data[$key]);	
				}

				// 格式化生育状况
				if ( $key == 'G4' ) {
						$data[$key] = preg_replace('/\s/', '', $data[$key]);	
						if ( strpos($data[$key], '√未育') !== false || 
							 strpos($data[$key], 'R未育') !== false || 
							 strpos($data[$key], '☑未育') !== false) {
							 $data[$key] = '未育';	
						}

						else if ( strpos($data[$key], '√已育') !== false || 
							 strpos($data[$key], 'R已育') !== false ||
							 strpos($data[$key], '☑已育') !== false ) {
							 $data[$key] = '已育';	
						}
				}

				// 格式化紧急联系人
				if ( $key == 'G11' ) {
					$data[$key] = preg_replace('/姓名|电话|:|：/', '', $data[$key]);	
					$data[$key] = preg_replace('/\s+/', ' ', $data[$key]);	
					$data[$key] = trim($data[$key]);
				}

				// 格式化婚姻状况
				if ( $key == 'G3' ) {
					$data[$key] = preg_replace('/\s/', '', $data[$key]);
					if ( strpos($data[$key], '√已') !== false ||
						 strpos($data[$key], 'R已') !== false ||
						 strpos($data[$key], '☑已') !== false ) {
						$data[$key] = '已婚';		
					}
					else if ( strpos($data[$key], '√未') !== false ||
						 strpos($data[$key], 'R未') !== false ||
						 strpos($data[$key], '☑未') !== false ) {
						$data[$key] = '未婚';		
					}
					else if ( strpos($data[$key], '√离异') !== false ||
						 strpos($data[$key], 'R离异') !== false ||
						 strpos($data[$key], '☑离异') !== false ) {
						$data[$key] = '离异';		
					}
				}

				// 格式化户口性质
				if ( $key == 'G5' ) {
					$data[$key] = preg_replace('/\s/', '', $data[$key]);
					if ( strpos($data[$key], '√城镇') !== false ||
						 strpos($data[$key], 'R城镇') !== false ||
						 strpos($data[$key], '☑城镇') !== false ) {
						$data[$key] = '城镇';		
					}
					else if ( strpos($data[$key], '√非城镇') !== false ||
						 strpos($data[$key], 'R非城镇') !== false ||
						 strpos($data[$key], '☑非城镇') !== false ) {
						$data[$key] = '非城镇';		
					}
				}

				// 格式化民族
				if ( $key == 'G2' ) {
					$data[$key] = preg_replace('/\s/', '', $data[$key]);
					if ( strpos($data[$key], '√汉') !== false ||
						 strpos($data[$key], 'R汉') !== false ||
						 strpos($data[$key], '☑汉') !== false ) {
						$data[$key] = '汉';		
					}
					else if ( strpos($data[$key], '√其他') !== false ||
						strpos($data[$key], 'R其他') !== false ||
						strpos($data[$key], '☑其他') !== false ) {
						$data[$key] = Utils::betweenCutStr('（', '）', $data[$key]);		
					}
				}

				// 格式化政治面貌
				if ( $key == 'I2' ) {
					$data[$key] = preg_replace('/\s/', '', $data[$key]);
					
					if ( strpos($data[$key], '□党员□团员□群众') !== false ) {
						 $data[$key] = '群众';		
					}

					else if ( strpos($data[$key], '√党员') !== false ||
						 strpos($data[$key], 'R党员') !== false ||
						 strpos($data[$key], '☑党员') !== false ) {
						 $data[$key] = '党员';		
					}
					else if ( strpos($data[$key], '√团员') !== false ||
						 strpos($data[$key], 'R团员') !== false ||
						 strpos($data[$key], '☑团员') !== false ) {
						 $data[$key] = '团员';		
					}
					else if ( strpos($data[$key], '√群众') !== false ||
						 strpos($data[$key], 'R群众') !== false ||
						 strpos($data[$key], '☑群众') !== false ) {
						 $data[$key] = '群众';		
					}
				}

				// 格式化B15-B18
				$data['B15'] = "*户口地址\n同户口薄";
				$data['B16'] = "*家庭地址 个人房\n产权/父母家地址";
				$data['B17'] = "*现居住地址 必须为\n工作城市居住地址";
				$data['B18'] = "*通讯地址 有效送\n达地址或家庭地址";
			}	
			return $data;
		}
}
