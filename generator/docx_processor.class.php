<?

// Copyright © 2015 Elena Bakun Contacts: <floppox@gmail.com>
// License: http://opensource.org/licenses/MIT

class docx_processor {

	private $document;
	private $original_tags;

	function docx_processor ($document) {
		$this-> processor($document);
		}

	function processor($document){
		$this->document = $document;
		}

	private function read_main_content(){
		if(empty($this->main_content))
			$this->main_content = $this->document->opened->getFromName('word/document.xml');
		}
	
	public function search_filds() {
		$this->read_main_content();

		$simple_tags = array();
		preg_match_all ( "|\{(.*)\}|U", $this->main_content, $simple_tags );

		$iterative_tags = array();
		preg_match_all ( "|\[\*(.*)\*\]|U", $this->main_content, $iterative_tags );

		$group_tags = array();
		preg_match_all ( "|\[\[(.*)\]\]|U", $this->main_content, $group_tags );

		$original_tags =  array();

		
		foreach ($simple_tags[1] as $k =>$t) {
			$original_tags['simple'][strip_tags($t)][]=$t;
			$simple_tags[1][$k] = strip_tags($t);
			}
		foreach ($iterative_tags[1] as $k =>$t) {
			$original_tags['iterative_tags'][strip_tags($t)][]=$t;
			$iterative_tags[1][$k]  = strip_tags($t);
			}
		foreach ($group_tags[1] as $k =>$t) {
			$original_tags['group'][strip_tags($t)][]=$t;
			$group_tags[1][$k]  = strip_tags($t);
			}
		$this->original_tags = (object)$original_tags;
		

		$tags =  new stdClass();
		$tags->simple = (object)$simple_tags[1];
		$tags->iterative = (object)$iterative_tags[1];
		$tags->group = (object)$group_tags[1];

		$this->tags = (object)$tags;
		
		
		
		return $tags;
		}
	public function replace_simple( $matrix) {

		foreach ($matrix as $key=>$value) {
			while ($original_key = array_shift($this->original_tags->simple[$key])) {
		
			$this->main_content = str_replace('{'.$original_key.'}', $value, $this->main_content);
		
			}
			}
		
		$this->document->opened->deleteName ( 'word/document.xml'); 
		$this->document->opened->addFromString('word/document.xml', $this->main_content);
		
		}


	public function build_groups($attributes) {
		
		}

	public function fill_iterative($matrix) {

	
		}

	
}
?>
