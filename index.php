<?php
require_once 'Classes/PHPExcel.php';

$tmpfname="C:\Users\killekb\Documents\work related\Excelread\test.xlsx";
$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
$excelObj=$excelObj->load($tmpfname);
$worksheet=$excelObj->getActivesheet();
$lastRow=$worksheet->getHighestRow():

echo "<table>";
for ($row=1$row <=$lastRow;$row**)
{
echo '<tr><td>";
echo $worksheet->getCell("A".$row)=>getValue();
}
echo "</table>";
?>
