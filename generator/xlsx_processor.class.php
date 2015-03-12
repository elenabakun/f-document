<?
class xlsx_processor {

	private $document;
	private $shared_strings;

	function xlsx_processor ($document) {
		$this-> processor($document);
		}

	function processor($document){
		$this->document = $document;
		}

	private function read_shared_strings($compile_strings=false) {
		if(empty($this->shared_strings->content))
			$this->shared_strings->content = $this->document->opened->getFromName('xl/sharedStrings.xml');
		if(empty($this->shared_strings->xml))
            $this->shared_strings->xml =  simplexml_load_string($this->shared_strings->content);
		if($compile_strings){
			$xml = simplexml_load_string($this->shared_strings->content);
			$this->rebuild_shared_strings_array($xml);
			}
		}
	private function rebuild_shared_strings_array($xml) {
		
			$sharedStringsArr = array();
			foreach ($xml->children() as $key =>$item) {
				$sharedStringsArr[] = (string)$item->t;
				}
			$this->shared_strings->arr =  $sharedStringsArr;

		}
	public function search_filds() {
		$this->read_shared_strings();

		$simple_tags = array();
		preg_match_all ( "|\{(.*)\}|U", $this->shared_strings->content, $simple_tags );

		$iterative_tags = array();
		preg_match_all ( "|\[\*(.*)\*\]|U", $this->shared_strings->content, $iterative_tags );

		$group_tags = array();
		preg_match_all ( "|\[\[(.*)\]\]|U", $this->shared_strings->content, $group_tags );

		$tags =  new stdClass();
		$tags->simple = (object)$simple_tags[1];
		$tags->iterative = (object)$iterative_tags[1];
		$tags->group = (object)$group_tags[1];

		$tags = (object)$tags;
		return $tags;
		}

	public function replace_simple( $matrix) {
		
		foreach ($matrix as $key=>$value) {
			$this->shared_strings->content = str_replace('{'.$key.'}', $value, $this->shared_strings->content);
			}
		$this->document->opened->deleteName ( 'xl/sharedStrings.xml'); $this->document->opened->addFromString('xl/sharedStrings.xml', $this->shared_strings->content);
		 $this->shared_strings->xml =  simplexml_load_string($this->shared_strings->content);
		}
	private function find_multiple_rows($sheets) {		
		foreach($sheets as $sheet)
			foreach($sheet->cels as $row=>$cels){
				
				foreach ($cels as $cell){
					preg_match("|\[\*(.*)\*\]|U", $cell,$tag);
					if (!empty($tag[1]))
					$multiple_rows[$sheet->sheet][$row][]=$tag[1];
					}
				}
		return (object)$multiple_rows;
		}


	private function count_iteratives($multiple_rows,$matrix) {
		$add = 0;
		foreach ($multiple_rows as $sheet=>$rows)
			foreach ($rows as $row=>$tags) {
				$row_count = 0;
				foreach ($tags as $tag) {
					if ($row_count<count($matrix->$tag))
						$row_count = count($matrix->$tag);
				}
				$multiple_rows_count[$sheet][$row+$add]=$row_count-1;
				$add+=$row_count-1;
			}
	
		return $multiple_rows_count;
		}


	private function undraw_rows($sheet, $row, $count){
		foreach ($sheet->xml->sheetData->row as $row_to_move) {
			$old_index = ((int)$row_to_move->attributes()->r);
			if ($old_index > $row) {
				$row_to_move->attributes()->r = $row_to_move->attributes()->r +$count;
				foreach ($row_to_move as $child) {
					$attr = $child->attributes();
					$child->attributes()->r = str_replace($old_index,($old_index+$count),$child->attributes()->r);
					}
				}
			}			
		}
	private function rebuild_merge_cells($sheet, $row, $count){
		$mc_total = ($sheet->xml->mergeCells && $sheet->xml->mergeCells->attributes()->count)?(int)$sheet->xml->mergeCells->attributes()->count:0;

		for($mc=0;$mc<$mc_total;$mc++){
			$mergecell=$sheet->xml->mergeCells->mergeCell[$mc];
			
			preg_match_all("/[\d]+/", $mergecell->attributes()->ref,$rownums);

			if($rownums[0][0]>$row) 
				$mergecell->attributes()->ref = str_replace($rownums[0][0],($rownums[0][0]+$count), $mergecell->attributes()->ref);

			if($rownums[0][1]>$row || ($rownums[0][1] == $row && $rownums[0][0]<$row)) 
				$mergecell->attributes()->ref = str_replace($rownums[0][1],$rownums[0][1]+$count, $mergecell->attributes()->ref);	

			if ($rownums[0][0]==$row && $rownums[0][1]==$row) {
				for ($i=0; $i<$count; $i++){
					$newmergecell = $sheet->xml->mergeCells->addChild('mergeCell');
					$rownums[0][0]+1+$i;
					$newmergecell->addAttribute('ref', str_replace($rownums[0][0],($rownums[0][0]+1+$i), $mergecell->attributes()->ref));
					$newmergecell->attributes()->ref;
					$sheet->xml->mergeCells->attributes()->count = $sheet->xml->mergeCells->attributes()->count + 1;
					}
				}
			}	
		}
	private function find_row($sheetData,$r) {
		foreach ($sheetData->row as $row) {
			if ($row->attributes()->r==$r) {
				return $row;
				}
			}
		return false;
		}
	private function clone_group($sheet, $row, $count){
		$prev_row = $row_to_clone = $this->find_row($sheet->xml->sheetData,($row));
		$header_to_clone = $this->find_row($sheet->xml->sheetData,($row-1));
		$row_to_clone->addAttribute('fl','r-orig');
		$header_to_clone->addAttribute('fl','h-orig');

		for ($i=0; $i<$count; $i++){
			$h_clone = $sheet->xml->sheetData->addChild('row');
			$r_clone = $sheet->xml->sheetData->addChild('row');
			$header_index = (int)$header_to_clone->attributes()->r;
			$row_index = (int)$row_to_clone->attributes()->r;

			foreach($row_to_clone->attributes() as $attr=>$value){
				$r_clone->addAttribute($attr,$value);
				}
			$r_clone->attributes()->r = $row_index+2+$i*2;
			
			

			foreach($header_to_clone->attributes() as $attr=>$value){
				$h_clone->addAttribute($attr,$value);
				}
			$h_clone->attributes()->r = $row_index+1+$i*2;
			
			$this->clone_row_cells($header_to_clone,$h_clone,$row_index-1,$row_index+1+$i*2);
			$this->clone_row_cells($row_to_clone,$r_clone,$row_index,$row_index+2+$i*2);

			$r_clone->attributes()->fl='r'.$i;
			$h_clone->attributes()->fl='h'.$i;

			$this->simplexml_insert_after($r_clone, $prev_row);
			$this->simplexml_insert_after($h_clone, $prev_row);

			$prev_row = $r_clone;
			}
			
		}
	private function clone_row_cells($row_to_clone,$clone,$old_row_index,$new_row_index){

		foreach ($row_to_clone->c as $cell){

				$newcell = $clone->addChild('c');
				
				foreach($cell->attributes() as $attr=>$value ){
					$newcell->addAttribute($attr,$value);
					}
				$newcell->attributes()->r = str_replace($old_row_index,$new_row_index,$cell->attributes()->r);

				if(!empty($cell->v)) {
					if($cell->attributes()->t!=s) {
						$newcell->addChild('v',$cell->v);
						}
					else {
						$v = (int)$cell->v;
						$newcell->addChild('v',count($this->shared_strings->xml->si));
						
						$newsi=$this->shared_strings->xml->addChild('si');
						$newsi->addChild('t',$this->shared_strings->xml->si[$v]->t);

						
						}
					}
				
				}
			
		}
	private function clone_iterative($sheet, $row, $count){
		
		$prev_row = $row_to_clone = $this->find_row($sheet->xml->sheetData,($row));
		for ($i=0; $i<$count; $i++){
			
			$clone = $sheet->xml->sheetData->addChild('row');
			$row_index = (int)$row_to_clone->attributes()->r;
			
			foreach($row_to_clone->attributes() as $attr=>$value){
				$clone->addAttribute($attr,$value);
				}
			$clone->attributes()->r = $row_index+1+$i;
			

			$this->clone_row_cells($row_to_clone,$clone,$row_index,$row_index+1+$i);
$this->rebuild_shared_strings_array($this->shared_strings->xml);
			
			$this->simplexml_insert_after($clone, $prev_row);
			$prev_row = $clone;
			}				
		}
	
	function simplexml_insert_after(SimpleXMLElement $insert, SimpleXMLElement $target){
		$target_dom = dom_import_simplexml($target);
		$insert_dom = $target_dom->ownerDocument->importNode(dom_import_simplexml($insert), true);
	
		if ($target_dom->nextSibling->nextSibling) {
			return $target_dom->parentNode->insertBefore($insert_dom, $target_dom->nextSibling);
			} 
		else {
			return $target_dom->parentNode->appendChild($insert_dom);
			}
		}

	private function find_groups($sheets) {
		foreach($sheets as $sheet)
			foreach($sheet->cels as $row=>$cels)
				foreach ($cels as $cell){
					preg_match("|\[\[(.*)\]\]|U", $cell,$tag);
					if (!empty($tag[1]))
						return  array('sheet'=>$sheet->sheet,'row'=>$row);
					}
		return false;
		
		}

	private function get_sheet($s) {
		foreach ($this->get_sheets() as $sheet){
			if ($sheet->sheet == $s)
				return $sheet;
			}
		return false;
		}

	public function build_groups($attributes) {
		
		$matrix=  $attributes[0];
		$map = $attributes[1];
		$this->read_shared_strings(true);

		//найти заголовок групп
		$group_title = $this->find_groups($this->get_sheets());
		

		//размножить группы вместе со следующим рядком в соответствии с картой 
		$count = count((array)$map)-1;
		$sheet= $this->get_sheet($group_title['sheet']);
		$row = $group_title['row'];
		$this->undraw_rows($sheet, $row+1, $count*2);
		$this->clone_group($sheet, $row+1, $count);
		$this->rebuild_merge_cells($sheet, $row+1, $count*2);
		$this->rebuild_hyperlinks($sheet, $row+1, $count, 2);
		
		$sheet->cels = $this->build_cells_array($sheet->xml);
		$this->document->opened->deleteName ( $sheet->sheet);
		$addSucess = $this->document->opened->addFromString( $sheet->sheet, $sheet->xml->asXML());

		//заменить тэги в заголовках групп значениями
		$this->shared_strings->content = $this->shared_strings->xml->asXML();
		
		//Здесь и ниже убрать это безобразие в отдельный метод 
		foreach($matrix as $key=>$values) {
			if (is_array($values)){
				foreach($values as $value){
					 $pos = strpos($this->shared_strings->content, '[['.$key.']]'); 
						if( $pos!==false )
							$this->shared_strings->content = substr_replace($this->shared_strings->content, $value, $pos, strlen('[['.$key.']]'));
					}
				$this->shared_strings->content = str_replace('[['.$key.']]','',$this->shared_strings->content);
				}
			else {
				$value = (string)$values;
				$this->shared_strings->content = str_replace('[['.$key.']]',$value,$this->shared_strings->content);
				}
			}
		for($i=0;$i<=$count;$i++)	{
			foreach($sheet->cels[$row+1] as $key){
				$value = str_replace('[*','[*g'.$i.'_',$key);
				$pos = strpos($this->shared_strings->content, $key); 
					if( $pos!==false )
						$this->shared_strings->content = substr_replace($this->shared_strings->content, $value, $pos, strlen($key));
				}
			}
		
		
		
		
		$this->shared_strings->xml = simplexml_load_string($this->shared_strings->content);
		$this->rebuild_shared_strings_array($this->shared_strings->xml);
		$sheet->cels = $this->build_cells_array($sheet->xml);
		
		
		
		$this->document->opened->deleteName ( 'xl/sharedStrings.xml'); $this->document->opened->addFromString('xl/sharedStrings.xml', $this->shared_strings->content);

		}
	
	private function rebuild_hyperlinks($sheet, $row, $count, $interval = 1) {
		if ($sheet->xml->hyperlinks->hyperlink){
			$R = $sheet->rels->xml;
			$h_total = count($sheet->xml->hyperlinks->hyperlink);
			for($h=0;$h<$h_total;$h++){
				
				$hl=$sheet->xml->hyperlinks->hyperlink[$h];
				preg_match_all("/[\d]+/", $hl->attributes()->ref,$rownums);
				if($rownums[0][0]>$row) {
					$hl->attributes()->ref = str_replace($rownums[0][0],($rownums[0][0]+$count), $hl->attributes()->ref);	
				}

			
				if ($rownums[0][0]==$row) {
					$original_rid = (string)$hl->attributes('r',1)->id;
					$hl->asXML();
					foreach ($R->Relationship as $rel) {
						if ($rel->attributes()->Id == $original_rid)
							$original_R = $rel;
						}
					for ($i=0; $i<$count; $i++){
						
						$newhl = $sheet->xml->hyperlinks->addChild('hyperlink');
						$newhl->addAttribute('ref', str_replace($rownums[0][0],($rownums[0][0]+$interval+$i*$interval), $hl->attributes()->ref));
						$rId = count($R->Relationship)+1;
						$newhl->addAttribute('xmlns:r:id', 'rId'.$rId);
						$newR = $R->addChild('Relationship');
						foreach ($original_R->attributes() as $attr_name => $attr_value){
							$newR->addAttribute($attr_name,$attr_value);
							}
						$newR->attributes()->Id = 'rId'.$rId;
						
						}
					}
				}
			$sheet->rels->content = $sheet->rels->xml->asXML();
			$sheet->rels->xml = simplexml_load_string($sheet->rels->content);
			$sheet->content = $sheet->xml->asXML();
			$sheet->xml = simplexml_load_string($sheet->content);
			}
		}

	public function fill_iterative($matrix) {

		$this->read_shared_strings(true);
		$multiple_rows = $this->find_multiple_rows($this->get_sheets());
		$multiple_rows_count = $this->count_iteratives($multiple_rows,$matrix);
	
		foreach ($this->get_sheets() as $sheet) {
			foreach ($multiple_rows_count[(string)($sheet->sheet)] as $row =>$count){
				$this->undraw_rows($sheet, $row, $count);
				$this->clone_iterative($sheet, $row, $count);
				$this->rebuild_merge_cells($sheet, $row, $count);
				$this->rebuild_hyperlinks($sheet, $row, $count);
				
				$sheet->cels = $this->build_cells_array($sheet->xml);

				$this->document->opened->deleteName ( $sheet->sheet);
				$addSucess = $this->document->opened->addFromString( $sheet->sheet, $sheet->xml->asXML());		
				}
			}
	
		$this->shared_strings->content = $this->shared_strings->xml->asXML();

		foreach($matrix as $key=>$values) {
			if (is_array($values)){
				foreach($values as $value){
					 $pos = strpos($this->shared_strings->content, '[*'.$key.'*]'); 
						if( $pos!==false )
							$this->shared_strings->content = substr_replace($this->shared_strings->content, $value, $pos, strlen('[*'.$key.'*]'));
					}
				$this->shared_strings->content = str_replace('[*'.$key.'*]','',$this->shared_strings->content);
				}
			else {
				$value = (string)$values;
				$this->shared_strings->content = str_replace('[*'.$key.'*]',$value,$this->shared_strings->content);
				}
			}

		if ($sheet->xml->hyperlinks->hyperlink){
			foreach ($sheet->xml->hyperlinks->hyperlink as $hl){
				preg_match_all("/[\d]+/", $hl->attributes()->ref,$rownums);
				
				$hl_order[$rownums[0][0]][]= array (
					'cell' => $hl->attributes()->ref,
					'rid' =>  $hl->attributes('r',1)->id
					);
				}
			ksort($hl_order);
			$order_index = 0;
			$hl_matrix = array();
			foreach ($hl_order as $hl_row) {
				foreach ($hl_row as $hl_cell) {
					$hl_matrix[(string)$hl_cell['rid']] = $order_index;
					}
				$order_index++;
				}
			
			foreach ($sheet->rels->xml as $rel){
				$rId = (string)$rel->attributes()->Id;
				

				preg_match_all ( "|\~(.*)\~|U", $rel->attributes()->Target, $simple_tags );
				preg_match_all ( "|\*\*(.*)\*\*|U", $rel->attributes()->Target, $iterative_tags );
				$simple_tags = $simple_tags[1];
				$iterative_tags = $iterative_tags[1];

				global $request;
				$fields = $request->matrix;

			
				foreach($simple_tags as $tag){
					$t = $fields->$tag;
					 $rel->attributes()->Target = str_replace('~'.$tag.'~',$t,$rel->attributes()->Target);
					}
				
				foreach($iterative_tags as $tag){
					$t = $fields->$tag;
					$i = (string)$hl_matrix[$rId];
					$rel->attributes()->Target = str_replace('**'.$tag.'**',$t[$i],$rel->attributes()->Target);
					}
				
				}
			$sheet->rels->content = $sheet->rels->xml->asXML();
			$this->document->opened->deleteName ( 'xl/worksheets/_rels/'.basename($sheet->sheet).'.rels'); 
			$this->document->opened->addFromString('xl/worksheets/_rels/'.basename($sheet->sheet).'.rels', $sheet->rels->content);
		}
		$this->document->opened->deleteName ( 'xl/sharedStrings.xml'); $this->document->opened->addFromString('xl/sharedStrings.xml', $this->shared_strings->content);

		
		}

	private $sheets;

	private function get_sheets(){
		if (empty($this->sheets)){

			for( $i = 0; $i < $this->document->opened->numFiles; $i++ ){ 
				$stat = $this->document->opened->statIndex( $i ); 
				if (strstr( $stat['name'],'xl/worksheets/')&&!strstr( $stat['name'],'_rels')){
					$s_content = '';
					$s_content =  $this->document->opened->getFromName($stat['name']);
					$xml = simplexml_load_string($s_content);

					$rels_content = 
						$this->document->opened->getFromName('xl/worksheets/_rels/'.basename($stat['name']).'.rels');
					$rels_xml =  simplexml_load_string($rels_content);

					$this->sheets[]=(object) array('sheet'=>$stat['name'], 'cels' => $this->build_cells_array($xml), 'xml'=>$xml, 'content' => $s_content, 'rels'=>(object)array('content'=>$rels_content,'xml'=>$rels_xml));

					
					
					}
				} 
			}
		
		return (object) $this->sheets;
		}
	private function build_cells_array($xml){
		$cels = array();
					foreach ($xml->sheetData->row as $row) 
						foreach ($row->c as $cell){
						preg_match("/[\d]+/", (string)$cell->attributes()->r,$r);
						$r = $r[0];
						$v = (string)$cell->v;
						
						if ($cell->attributes()->t == 's')
							$v=(string)$this->shared_strings->arr[$v];
						if ($v) 
							$cels[$r][]=$v;

						}
		return $cels;
		
		}
}
?>
