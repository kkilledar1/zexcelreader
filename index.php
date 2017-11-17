<?php
/** Include PHPExcel **/
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
//include "Classes/PHPExcel.php";
//require_once ('PHPExcel.php');
// Create new PHPExcel object
echo 'Hello World';

 $tmpfname = "test.xls";
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);
		$worksheet = $excelObj->getSheet(0);
		$lastRow = $worksheet->getHighestRow();
		$rowIterator = $objPHPExcel->getActiveSheet()->getRowIterator()
			
$foundInCells = array();
$searchValue = 'B';
		echo "<table>";
		for ($row = 1; $row <= $lastRow; $row++) 
		{
			 echo "<tr><td>";
			 echo $worksheet->getCell('A'.$row)->getValue();
			 echo "</td><td>";
			 echo $worksheet->getCell('B'.$row)->getValue();
			 echo "</td><tr>";
			if ($cell->getValue("A'.$row) == $searchValue)
			{
                        echo 'Found it';
                         }
		}
		echo "</table>";
                echo 'Excel read';


?>
