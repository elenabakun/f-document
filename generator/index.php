<?

// Copyright © 2015 Elena Bakun Contacts: <floppox@gmail.com>
// License: http://opensource.org/licenses/MIT

function __autoload($classname) {
    include_once("./".strtolower($classname) .".class.php");
}

$request = new Request();

$document = new Document($request->template);

$document->process();

$matrix = new Compiler($document->search_filds(),$request->matrix);

$document->replace_simple($matrix->simple); 

if(!empty($matrix->groups_map)){
	$document-> build_groups($matrix->group, $matrix->groups_map);
	$matrix->expand_iteratives();
	}

$document->fill_iterative($matrix->iterative); 

if($request->extention){
	$action = 'extention';
	$extention = $request->extention;
	}
elseif($request->action){
	$action = $request->action;
	$extention = false;
	}
else {
	$action = 'show_link';
	$extention = false;
	}
$document->close($action,$extention);

?>
