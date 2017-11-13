<?php
include 'PHPExcel.php'

$inputFileName="C:\Users\killekb\Documents\work related\Excelread\test.xlsx";
$excelReader = PHPExcel_IOFactory::createReaderForFile($inputFileName);
$excelObj=$excelObj->load($tmpfname);
$worksheet=$excelObj->getActivesheet();
$lastRow=$worksheet->getHighestRow();

echo "Hello World Excel";

?>
