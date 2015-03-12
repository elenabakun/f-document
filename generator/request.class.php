<?

// Copyright © 2015 Elena Bakun Contacts: <floppox@gmail.com>
// License: http://opensource.org/licenses/MIT

class Request {
	
	public $template;
	
	public $matrix;

	function Request() {
		$this->matrix  = new stdClass();
		$this->read_request(); 
		$this->arrange_request();
		}

	private $request_data = array();
	
	private $request_method;
	private $request_encoding;
	public $action;
	public $extention;

	private function read_request() {
		$this->request_method = $_SERVER['REQUEST_METHOD'];
		$this->request_data = $_REQUEST;
		$this->request_encoding = $_REQUEST['_encoding'];
		}

	private function arrange_request() {
		foreach ($this->request_data as $key=>$value) {
			$key=$this->clean_request($key);

			if ($key=='_template') $this->template = $this->clean_request($value); 

			elseif ($key=='_encoding'){}
			elseif ($key=='_action')$this->action = $this->clean_request($value);
			elseif ($key=='_extention')$this->extention = $this->clean_request($value);

			elseif (is_array($value)){
				$k = array();
				foreach ($value as $v){
					 $k[] = $this->clean_request($v);
					}
				$this->matrix->$key = $k;
				}
			else $this->matrix->$key = $this->clean_request($value);
		}
		
		}
	private function clean_request($value) {
		if ($this->request_method == "GET") 
			$value = urldecode($value);

		if ($this->request_encoding == "windows-1251") 
			$value=iconv("windows-1251", "utf-8",$value);


		return $value;
		}

	}


?>
