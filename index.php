<?php
$method = $_SERVER['REQUEST_METHOD'];
 
// Process only when method is POST
//if($method == 'POST'){
    $requestBody = file_get_contents('php://input');
    $json = json_decode($requestBody);
 
    $searchValue= $json->result->parameters->affloc;
 

/** Include PHPExcel **/
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
//include "Classes/PHPExcel.php";
//require_once ('PHPExcel.php');
// Create new PHPExcel object
//echo 'Hello World';
//$searchValue = $affloc;

 $tmpfname = "GRC details.xls";
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
			$column =$worksheet->getcell('B'.$row)->getValue();
			$speech= "Your security coordinator is $column";	
				break;
                         }
		}	 
			 //echo $column;

if (empty($column)){
	$speech="No results from search.Please contact Helpdesk";
	//echo "No results";
}
//echo "</table>";
		//echo 'Excel read';
 
    $response = new \stdClass();
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
