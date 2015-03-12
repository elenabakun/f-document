<?
class  Document{

	public function process() {
		$class = $this->file_type.'_processor';
		$this->processor = new $class($this);
		}

	public function __call($name, $arguments) {
		if (count($arguments) == 1 ) $arguments = $arguments[0];
		return $this->processor->$name($arguments);
		}

	function Document($template_name, $template_path = 'templates/', $target_path='downloads/') {
		$this->template_name = $template_name;
		$this->open_archive($this->make_work_copy($template_path,$target_path));
		$this->file_type = $this->check_document_type();
		}

	private $template_name;
	private $template_file;
	public $temporary_name;
	public $opened; 
	private $file_type;

	private function make_work_copy($template_path,$target_path) {
		$this->temporary_name = $this->unick_name($this->template_name);
		$template_file = $_SERVER['DOCUMENT_ROOT'].$template_path.$this->template_name;
		$temporary_file = $_SERVER['DOCUMENT_ROOT'].$target_path.$this->temporary_name;
		copy ($template_file, $temporary_file);
		return $temporary_file;
		}
	private function unick_name($template_name){
		return time().$template_name;
		}
	private function open_archive($file_name) {
		$this->opened = new ZipArchive;
		$this->opened->open($file_name);
		}
	private function check_document_type() {
		if ($this->opened->locateName('xl/sharedStrings.xml')!==false) return 'xlsx';
		if ($this->opened->locateName('word/document.xml')!==false) return 'docx';
		else return false;
		}

	public function close($action = false,$extention = '') {
		$this->opened->close();
		if ($action == 'extention') {
			$this->special_action($extention);
			}
		else if ($action) 	{
			$this->$action();
			}
			
		}
	
	function file_download() {
		$file = '../downloads/'.$this->temporary_name;
		if (file_exists($file)) {
			if (ob_get_level()) {
			  ob_end_clean();
			}
			header('Content-Description: File Transfer');
			header('Content-Type: application/octet-stream');
			header('Content-Disposition: attachment; filename=' . basename($this->template_name));
			header('Content-Transfer-Encoding: binary');
			header('Expires: 0');
			header('Cache-Control: must-revalidate');
			header('Pragma: public');
			header('Content-Length: ' . filesize($file));
			readfile($file);
			unlink($file);
			}
		}
	function show_link(){
		$file = '../downloads/'.$this->temporary_name;
		if (file_exists($file)) {
			$link = 'http://'.$_SERVER['SERVER_NAME'].'/downloads/'.$this->temporary_name;
			echo '<a href="'.$link.'">Скачать '.$this->template_name.'</a>';
			}
		}
	function show_direct_link(){
		$file = '../downloads/'.$this->temporary_name;
		if (file_exists($file)) {
			echo 'http://'.$_SERVER['SERVER_NAME'].'/downloads/'.$this->temporary_name;
			}
		}
	function special_action($extention) {
		if (file_exists('../extentions/'.$extention.'.php'))
			include '../extentions/'.$extention.'.php';
		else $this->show_link();
	
		}

}
?>
