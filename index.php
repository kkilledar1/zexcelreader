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
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);
//$objReader->setLoadSheetsOnly($sheetname)
		$worksheet = $excelObj->getSheet(0);
		$lastRow = $worksheet->getHighestRow();
		
//echo "<table>";
		for ($row = 1; $row <= $lastRow; $row++) {
			 //echo "<tr><td>";
			 //echo $worksheet->getCell('A'.$row)->getValue();
			 //echo "</td><td>";
			 //echo $worksheet->getCell('B'.$row)->getValue();
			 //echo "</td><tr>";
		$compare=$worksheet->getCell('A'.$row)->getValue();
			
			if ($compare == $searchValue)	{
			$column =$worksheet->getcell('J'.$row)->getValue();
			//$speech= "CIR status is $column";
				$CIRlead=$worksheet->getcell('L'.$row)->getValue();
				$projecttype=$worksheet->getcell('I'.$row)->getValue();
				$processarea=$worksheet->getcell('D'.$row)->getValue();
				$priority=$worksheet->getcell('M'.$row)->getValue();
				break;
                         }
		}	 
			 //echo $column;

if (empty($column)){
	$speech="No results from search.Please contact CIR Helpdesk";
	//echo "No results";
}
//echo "</table>";
		//echo 'Excel read';
 
    $response = new \stdClass();

if ($action=="findcirrecord") {
	$speech= "CIR status is $column";

elseif ($action=="findcirlead") {
	$speech =" CIR lead is $CIRlead";
}
}
elseif ($action=="findprocarea"){
	$speech="This CIR is managed by $processarea";
}

    $response->speech = "$speech";
    $response->displayText = $speech;
    $response->source = "webhook";
    echo json_encode($response);
//}
//else
//{
  //  echo "Method not allowed";
//}
 
?>
