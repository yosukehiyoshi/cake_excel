<?php
/**
 * Object::DownloadComponent
 * Utility component for downloading
 *
 * @link          https://github.com/yosukehiyoshi/cake_excel
 * @package       cake_excel
 * @subpackage    controllers.components
 * @license       MIT License (http://www.opensource.org/licenses/mit-license.php)
 */
class DownloadComponent extends Object {

/**
 * Controller instance
 * @access public
 * @var Controller
 */
	public $Controller;

/**
 * MIME types
 * @access public
 * @var array
 */
	protected $_fileTypes = array(
		'Excel2007' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
	);

/**
 * Called before the Controller::beforeFilter().
 *
 * @param object $controller Controller with components to initialize
 * @return void
 * @access public
 * @link http://book.cakephp.org/view/998/MVC-Class-Access-Within-Components
 */
	public function initialize (Controller &$controller) {
		$this->Controller = $controller;
	}

/**
 * Set content type
 * @access public
 * @param string $type Key for $_fileTypes
 * @return boolean
 */
	public function setFileType ($type='') {
		if (!isset($this->_fileTypes[$type])) : return false; endif;
		$this->Controller->header("Content-Type: {$this->_fileTypes[$type]}");
		return true;
	}

/**
 * Set HTTP headers and file name
 * @access public
 * @param string $fileName File name for downloading
 * @return boolean
 */
	public function setDownloadHeaders ($fileName='') {
		if (!$fileName) : return false; endif;
		$this->Controller->header("Content-Disposition: attachment;filename=\"{$fileName}\"");
		$this->Controller->header('Cache-Control: max-age=0');
		// If IE9
		$this->Controller->header('Cache-Control: max-age=1');
		// If IE over SSL
		$this->Controller->header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
		$this->Controller->header('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
		$this->Controller->header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
		$this->Controller->header('Pragma: public'); // HTTP/1.0
		return true;
	}

}
