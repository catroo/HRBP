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

class Utils {

		 /**
		 * @格式化cellName 
		 *
		 * @param $ascii
		 *
		 * @return string 
		 */
		public static function formatCellName($ascii)
		{
				if ( $ascii <= 90 ) {
						return chr($ascii); 
				} else if ($ascii > 90 && $ascii < 117) {
						return 'A' .chr($ascii % 90 + 64);
				} else if ($ascii >= 117 && $ascii < 143) {
						return 'B' . chr(($ascii-26) % 90 + 64);
				} else if ($ascii >= 143) {
						return 'C' . chr(($ascii-26*2) % 90 + 64);
				}
		}

		/**
		* 写入错误日志 
		*
		* @param $error
		*
		* @return void 
		*/
		public static function errorLog($error)
		{
			$error = '[' . date('Y-m-d H:i:s') . ']' .$error . "\n";
			file_put_contents(dirname(dirname(__FILE__)) . '/logs/error.log', $error, FILE_APPEND);
			echo $error;
		}

		
		/**
		* 获取两个字符串中间的字符串 
		*
		* @param $begin
		* @param $end
		* @param $str
		*
		* @return 
		*/
		public static function betweenCutStr($begin,$end,$str){
                $b = mb_strpos($str,$begin) + mb_strlen($begin);
                $e = mb_strpos($str,$end) - $b;

                return mb_substr($str,$b,$e);
        }
}
