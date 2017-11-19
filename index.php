<?php
$method = $_SERVER['REQUEST_METHOD'];
 
// Process only when method is POST
if($method == 'POST'){
    $requestBody = file_get_contents('php://input');
    $json = json_decode($requestBody);
 
    $affloc = $json->result->parameters->affloc;
 

/** Include PHPExcel **/
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
//include "Classes/PHPExcel.php";
//require_once ('PHPExcel.php');
// Create new PHPExcel object
//echo 'Hello World';
$searchValue = $affloc;

 $tmpfname = "test.xls";
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);
		$worksheet = $excelObj->getSheet(0);
		$lastRow = $worksheet->getHighestRow();
		
echo "<table>";
		for ($row = 1; $row <= $lastRow; $row++) {
			 //echo "<tr><td>";
			 //echo $worksheet->getCell('A'.$row)->getValue();
			 //echo "</td><td>";
			 //echo $worksheet->getCell('B'.$row)->getValue();
			 //echo "</td><tr>";
		$compare=$worksheet->getCell('A'.$row)->getValue();
			
			if ($compare == $searchValue)	{
			$column =$worksheet->getcell('B'.$row)->getValue();
				
				break;
                         }
		}	 
			 echo $column;

if (empty($column)){
	echo "No results";
}
echo "</table>";
		//echo 'Excel read';
 
    $response = new \stdClass();
    $response->speech = $column;
    $response->displayText = $column;
    $response->source = "webhook";
    echo json_encode($response);
?>
