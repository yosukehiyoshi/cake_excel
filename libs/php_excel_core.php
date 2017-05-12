<?php
/**
 * PHPExcel extended class for CakePHP 1.3.
 *
 * @link          https://github.com/yosukehiyoshi/cake_excel
 * @package       cake_excel
 * @subpackage    libs
 * @license       MIT License (http://www.opensource.org/licenses/mit-license.php)
 */

App::import(
	null,
	'CakeExcel.PHPExcel',
	array(
		'search'    => App::pluginPath('CakeExcel'). 'vendors/PHPExcel/Classes',
		'file'      => 'PHPExcel.php'
	)
);

class PhpExcelCore extends PHPExcel {}
