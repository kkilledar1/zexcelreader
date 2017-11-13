<?php
require_once 'Classes/PHPexcel.php';

$tmpfname="C:\Users\killekb\Documents\work related\Excelread\test.xlsx";
$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
$excelObj=$excelObj->load($tmpfname);
$worksheet=$excelObj->getActivesheet();
$lastRow=$worksheet->getHighestRow();

echo "Hello World Excel";

?>
