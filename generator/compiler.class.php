<?
class Compiler {
	public $simple;
	public $iterative;
	public $group;
	public $groups_map;
    private $matrix;
	function compiler($fields, $request) {
		$this->build_matrix($fields, $request);
		if (!empty($this->group) && !empty($request->_group_key)){
			 $this->builg_groups_map($request);
			}
		}
	public function expand_iteratives(){
		foreach  ($this->groups_map as $g_index => $g_content){
			
			foreach ($g_content as $row_index){

				foreach ($this->iterative as $tag => $replaces) {
					$g_tag = 'g'.$g_index.'_'.$tag;
					$extendet_iteratives[$g_tag][] = $replaces[$row_index];

					}
				}
			
			}
		foreach ($extendet_iteratives as $tag => $replaces){
			$this->iterative->$tag = $replaces;
			}
			
		}

	private function build_matrix($fields, $request) {
		foreach ($fields->simple as $field) {
			$this->simple->$field = $request->$field;
			}
		foreach ($fields->iterative as $field) {
			$this->iterative->$field = $request->$field;
			}
		foreach ($fields->group as $field) {
			$g = 'g_'.$field;
			$this->group->$field = $request->$g;
			}		
		}
	
	
	private function builg_groups_map($request) {

		$group='';
		$g_key = $request->_group_key;
		foreach ($request->_group as $index=>$current) {
			if ($group!==$current) {
				$group = $current;
				$groups_map[$group] =  array();
				
				}
			foreach($this->iterative->$g_key as $key => $value) {
				if($key==$index) 
					$groups_map[$group][] = $key;
				}
			}
		$this->groups_map = (object)$groups_map;

		}
	
	}
?>
