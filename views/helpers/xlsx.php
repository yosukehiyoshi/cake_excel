<?php
/**
 * AppHelper::XlsxHelper
 *
 * @link          https://github.com/yosukehiyoshi/cake_excel
 * @package       cake_excel
 * @subpackage    views.helpers
 * @license       MIT License (http://www.opensource.org/licenses/mit-license.php)
 * @require       PHPExcel(https://github.com/PHPOffice/PHPExcel)
 */

App::import('Lib', array('CakeExcel.PhpExcelCore'));

class XlsxHelper extends AppHelper {

/**
 * View instance
 * @access public
 * @static
 * @var View
 */
	public static $View;

/**
 * Constructor
 * @access public
 */
	public function __construct () {
		parent::__construct();
		self::$View =& ClassRegistry::getObject('View');
	}

/**
 * Generate Spreadsheet file from template
 * @access public
 * @param string $filePath Template Spreadsheet file path
 * @param array &$data Array data for apply
 * @param string $type Spreadsheet type
 * @return void
 */
	public function generate ($filePath='', &$data, $type='Excel2007') {
		// Load template
		$XlsxReader = PHPExcel_IOFactory::createReader($type);
		$PhpExcel = $XlsxReader->load($filePath);
		// Apply data
		$this->_applyData($PhpExcel, $data);
		// Output
		$XlsxWriter = PHPExcel_IOFactory::createWriter($PhpExcel, $type);
		$XlsxWriter->save('php://output');
	}

/**
 * Apply data
 * @access public
 * @param PHPExcel &$PhpExcel PHPExcel instance
 * @param array &$data Array data for apply
 * @return PHPExcel
 */
	protected function _applyData (&$PhpExcel, &$data) {
		$workSheet = $PhpExcel->getActiveSheet();

		foreach ($workSheet->getCellCollection() as $cell) {
			$currentCell = $workSheet->getCell($cell);
			$cellValue = $currentCell->getValue();
			if (!$cellValue) : continue; endif;
			// Parse Mustache template
			if (!preg_match_all('/{{([^{}]+)}}/', $cellValue, $match)) : continue; endif;
			$names = $match[1];
			foreach ($names as &$name) {
				$methodName = $name;
				$arguments = array();

				if (preg_match('/^(.+)\((.+)\)$/', $name, $match)) {
					$methodName = $match[1];
					foreach (explode(',', $match[2]) as $argument) {
						$arguments[] = trim($argument);
					}
				}

				if (method_exists($this, $methodName)) {
					$tmp = call_user_func_array(
						array($this, $methodName),
						array($data, $currentCell, $arguments)
					);
					$cellValue = str_replace("{{{$name}}}", $tmp, $cellValue);
					continue;
				}

				if (array_key_exists($methodName, $data)) {
					$cellValue = str_replace("{{{$name}}}", $data[$methodName], $cellValue);
					continue;
				}

				$this->log("Missed: {$name}", LOG_DEBUG);
			}
			$currentCell->setValue(str_replace("\n", PHP_EOL, $cellValue));
		}

		return $PhpExcel;
	}

}
