<?php
$method = $_SERVER['REQUEST_METHOD'];
 
// Process only when method is POST
//if($method == 'POST'){
    $requestBody = file_get_contents('php://input');
    $json = json_decode($requestBody);
 
    $searchValue= $json->result->parameters->cirno;
$action=$json->result->action;
 

/** Include PHPExcel **/
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
//include "Classes/PHPExcel.php";
//require_once ('PHPExcel.php');
// Create new PHPExcel object
//echo 'Hello World';
//$searchValue = $affloc;

$tmpfname = "CIR tracker.xls";
//$tmpfname = "PEG.xls";
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);

		$worksheet = $excelObj->getSheet(0);
		$lastRow = $worksheet->getHighestRow();
		
		for ($row = 1; $row <= $lastRow; $row++) {
			 
			$compare=$worksheet->getCell('A'.$row)->getValue();
			
			if ($compare == $searchValue)	{
				$column =$worksheet->getcell('J'.$row)->getValue();
				$CIRlead=$worksheet->getcell('L'.$row)->getValue();
				$projecttype=$worksheet->getcell('I'.$row)->getValue();
				$processarea=$worksheet->getcell('D'.$row)->getValue();
				$priority=$worksheet->getcell('M'.$row)->getValue();
				//$Filter_out[$row]=$column;
				break;
                                                         }
		                                          }	 
			
//if (empty($column)){
//	$speech="No results from search.Please contact CIR Helpdesk";
	  //         }
 
    $response = new \stdClass();
if (empty($column)){
	$speech="No results from search.Please contact CIR Helpdesk";
                   }
	
	elseif ($action=="findcirrecord") {
		$speech= "CIR status is $column";
	              			}
	
        elseif ($action=="findcirlead") {
	$speech =" CIR lead is $CIRlead";
	                                }

	elseif ($action=="findprocarea"){
	$speech="This CIR is managed by $processarea";
	                                 }
//$tmp_column=implode (" ",$Filter_out);	
//$Filter_out = array ();
//$Filter_out = array ('1','2','3');
   $response->speech = "$speech";
//$response->speech ="$tmp_column";
  $response->displayText = $speech;
 //$response->displayText= $tmp_column;
    $response->source = "webhook";
    echo json_encode($response);

 
?>
