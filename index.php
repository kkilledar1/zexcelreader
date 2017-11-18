<?php
/** Include PHPExcel **/
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
//include "Classes/PHPExcel.php";
//require_once ('PHPExcel.php');
// Create new PHPExcel object
//echo 'Hello World';
$searchValue = 'B';

 $tmpfname = "test.xls";
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);
		$worksheet = $excelObj->getSheet(0);
		$lastRow = $worksheet->getHighestRow();
		
echo "<table>";
		for ($row = 2; $row <= $lastRow; $row++) {
			 //echo "<tr><td>";
			 //echo $worksheet->getCell('A'.$row)->getValue();
			 //echo "</td><td>";
			 //echo $worksheet->getCell('B'.$row)->getValue();
			 //echo "</td><tr>";
		$compare=$worksheet->getCell('A'.$row)->getValue();
			$column =$worksheet->getcell('B'.$row)->getValue();
			if ($compare == $searchValue)	{
			echo $column;
				break;
                         }
		}	 
			 
echo "</table>";
		//echo 'Excel read';

?>
