<?php
/*********************************************
    PRICE SHEET GATHERING INTERFACE
    BY: Brandon Jensen
    For: Material Flow & Conveyor Systems INC.
    Copyright 2018, All Rights Reserved
    price.php
**********************************************/
// Increase php runtime limit for large files, may not be necessary, chrome was the cultprit,
// giving up and timing out too quicking on one very large company.
set_time_limit(1000);
// Composer autoload
require 'vendor/autoload.php';
// Require database setup variables
require 'sql_config.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Initiate database connection
$conn = mysqli_connect($local, $user, $passwerd, $db);

// Check to see if postback happened yet
if(empty($_POST['manufacturer'])){
  // If not, have user select a manufacturer
  if(!$conn) {//check database connection
	// Database connection was not successful
	  $realIP = file_get_contents("http://ipecho.net/plain");	
	  die('Could not connect to the server! ' . mysqli_connect_error() . "<br>Server IP: " . $local . "<br>Actual IP: " . $realIP);  
  } else {
    // Database connection was successful
    $sql = "SELECT manufactName FROM manufacturer";
    // Retrieve Manufacturer Names
    $result = mysqli_query($conn, $sql);
    echo "<h1>SELECT MANUFACTURER</h1>";
    if (mysqli_num_rows($result)>0){
    	echo "<form action='Price.php' method='POST'><select name='manufacturer'>";
    	// Output Manufacturer Names
    	while($row = mysqli_fetch_assoc($result)) {
    		echo "<option value='".$row['manufactName']."'>".$row['manufactName']."</option>";
    	}
        echo "</select><button type='submit'>GET PRICES</button>";
        echo "<br><input type='checkbox' name='products' value='true' checked> Include Products<br><input type='checkbox' name='options' value='true' checked> Include Options </form>";
    } else { echo "0 Results Found"; }
  }
} elseif ($_POST['products'] || $_POST['options']) {
	$sheetrow = 3;
	$pricerow = 1;
	
	// Create New Sheet
	$spreadsheet = new Spreadsheet();
	$sheet = $spreadsheet->getActiveSheet(); 
	$sheet->setTitle($_POST['manufacturer']);
	$ourprices = $spreadsheet->createSheet();
	$ourprices->setTitle('Our_Prices');
	$msrp = $spreadsheet->createSheet();
	$msrp->setTitle('MSRP_Cost');

	// Create Competitor Pricing Sheets
	for ($i=1; $i <= 12 ; $i++) { 
		$comp = $spreadsheet->createSheet();
		$comp->setTitle('COMP' . $i);
		$comp->setCellValue('A1', 'LINK');
		$comp->setCellValue('B1', 'ITEM');
		$comp->setCellValue('C1', 'PRICE');
	}
	// Cell Style Array
	$stylearray = [
		'borders' => [
			'allBorders' => [
				'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				'color' => ['argb' => '00000000'],
			],
		],
	];

	// Set Default Display Setting
	$spreadsheet->getDefaultStyle()->getFont()->setName('Trebuchet MS');
	$spreadsheet->getDefaultStyle()->getFont()->setSize(10);
	$sheet->getPageSetup()->setOrientation(\PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE);
	$sheet->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 2);
	$sheet->freezePane('A3');
  
	// Set Price List Title
	$sheet->getStyle('A1')->getFont()->setSize(20);
	$sheet->getStyle("A1:W2")->getFont()->setBold(true);
	$sheet->getStyle("A2:W2")->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
	$sheet->getStyle("A1:W2")->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('bfbfbf');
	$sheet->setCellValue('A1', $_POST['manufacturer'] . ' PRICE LIST');
  
	// Column Titles
	$sheet->getStyle('A2:W2')->getAlignment()->setWrapText(true);
	$sheet->setCellValue('A2', 'MODEL')
		->setCellValue('B2', 'OUR PRICE')
		->setCellValue('C2', 'LIST PRICE')
		->setCellValue('D2', 'OUR COST')
		->setCellValue('E2', 'COMPETITOR')
		->setCellValue('F2', 'NEW PRICE')
		->setCellValue('G2', 'COMP - PRICE')
		->setCellValue('H2', 'COMP / COST')
		->setCellValue('I2', 'COMP NAME')
		->setCellValue('J2', 'OLD MARGIN')
		->setCellValue('K2', 'NEW MARGIN')
		->setCellValue('L2', 'COMP1')
		->setCellValue('M2', 'COMP2')
		->setCellValue('N2', 'COMP3')
		->setCellValue('O2', 'COMP4')
		->setCellValue('P2', 'COMP5')
		->setCellValue('Q2', 'COMP6')
		->setCellValue('R2', 'COMP7')
		->setCellValue('S2', 'COMP8')
		->setCellValue('T2', 'COMP9')
		->setCellValue('U2', 'COMP10')
		->setCellValue('V2', 'COMP11')
		->setCellValue('W2', 'COMP12');

	$ourprices->setCellValue('A1', 'ITEM')
		->setCellValue('B1', 'PRICE');

	$msrp->setCellValue('A1', 'ITEM')
		->setCellValue('B1', 'MSRP')
		->setCellValue('C1', 'COST');
	
	function addItem ($ItemName) {
		global $sheetrow;
		global $sheet;
		$sheetrow++;
		$sheet->getStyle('A'. $sheetrow)->getFont()->setSize(12);
		$sheet->getStyle('A' . $sheetrow)->getFont()->setBold(true);
		$sheet->setCellValue('A' . $sheetrow, $ItemName);
		$sheetrow++;		
	}
	
	function addModel ($modelname, $modelPrice) {
		global $sheetrow;
		global $pricerow;
		global $sheet;
		global $ourprices;
		global $stylearray;
		$sheet->getStyle('A' . $sheetrow . ':K' . $sheetrow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
		$sheet->getStyle('A' . $sheetrow . ':K' . $sheetrow)->applyFromArray($stylearray);
		$sheet->setCellValue('A' . $sheetrow, $modelname);
		$ourprices->setCellValue('A' . $pricerow, $modelname)
			->setCellValue('B' . $pricerow, $modelPrice);
		$sheet->setCellValue('B' . $sheetrow, '=VLOOKUP(A' . $sheetrow . ',Our_Prices!A:B, 2, FALSE)')
			->setCellValue('C' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',MSRP_Cost!A:C,2,FALSE)),VLOOKUP(A' . $sheetrow . ',MSRP_Cost!A:C,2,FALSE),"")')
			->setCellValue('D' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',MSRP_Cost!A:C,3,FALSE)), VLOOKUP(A' . $sheetrow . ',MSRP_Cost!A:C,3,FALSE),"")')
			->setCellValue('E' . $sheetrow, '=IF(MIN(L' . $sheetrow . ':W' . $sheetrow . ')>0, MIN(L' . $sheetrow . ':W' . $sheetrow . '),"")')
			->setCellValue('G' . $sheetrow, '=IFERROR(E' . $sheetrow . '-B' . $sheetrow . ', "")')
			->setCellValue('H' . $sheetrow, '=IFERROR(E' . $sheetrow . '/D' . $sheetrow . ', "")')
			->setCellValue('I' . $sheetrow, '=IF(ISNUMBER($E' . $sheetrow . '), IF($E' . $sheetrow . '=$L' . $sheetrow . ',$L$2,IF($E' . $sheetrow . '=$M' . $sheetrow . ',$M$2,IF($E' . $sheetrow . '=$N' . $sheetrow . ',$N$2,IF($E' . $sheetrow . '=$O' . $sheetrow . ',$O$2,IF($E' . $sheetrow . '=$P' . $sheetrow . ',$P$2, IF($E' . $sheetrow . '=$Q' . $sheetrow . ',$Q$2, IF($E' . $sheetrow . '=$R' . $sheetrow . ',$R$2, IF($E' . $sheetrow . '=$S' . $sheetrow . ',$S$2, IF($E' . $sheetrow . '=$T' . $sheetrow . ',$T$2, IF($E' . $sheetrow . '=$U' . $sheetrow . ',$U$2, IF($E' . $sheetrow . '=$V' . $sheetrow . ',$V$2, IF($E' . $sheetrow . '=$W' . $sheetrow . ',$W$2, "Not Found") ) ) ) ) ) ) ) ) ) ) ), "Not Found")')
			->setCellValue('J' . $sheetrow, '=IF(D' . $sheetrow . ' > 0,IFERROR((B' . $sheetrow . '-D' . $sheetrow . ')/B' . $sheetrow . ',""),"")')
			->setCellValue('K' . $sheetrow, '=IF(D' . $sheetrow . ' > 0,IFERROR((F' . $sheetrow . '-D' . $sheetrow . ')/F' . $sheetrow . ',""),"")')
			->setCellValue('L' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP1!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP1!B:C,2,FALSE),"")')
			->setCellValue('M' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP2!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP2!B:C,2,FALSE),"")')
			->setCellValue('N' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP3!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP3!B:C,2,FALSE),"")')
			->setCellValue('O' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP4!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP4!B:C,2,FALSE),"")')
			->setCellValue('P' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP5!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP5!B:C,2,FALSE),"")')
			->setCellValue('Q' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP6!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP6!B:C,2,FALSE),"")')
			->setCellValue('R' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP7!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP7!B:C,2,FALSE),"")')
			->setCellValue('S' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP8!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP8!B:C,2,FALSE),"")')
			->setCellValue('T' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP9!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP9!B:C,2,FALSE),"")')
			->setCellValue('U' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP10!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP10!B:C,2,FALSE),"")')
			->setCellValue('V' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP11!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP11!B:C,2,FALSE),"")')
			->setCellValue('W' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP12!B:C,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP12!B:C,2,FALSE),"")')
			->setCellValue('AA' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP1!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP1!E:F,2,FALSE),"")')
			->setCellValue('AB' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP2!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP2!E:F,2,FALSE),"")')
			->setCellValue('AC' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP3!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP3!E:F,2,FALSE),"")')
			->setCellValue('AD' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP4!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP4!E:F,2,FALSE),"")')
			->setCellValue('AE' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP5!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP5!E:F,2,FALSE),"")')
			->setCellValue('AF' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP6!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP6!E:F,2,FALSE),"")')
			->setCellValue('AG' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP7!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP7!E:F,2,FALSE),"")')
			->setCellValue('AH' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP8!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP8!E:F,2,FALSE),"")')
			->setCellValue('AI' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP9!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP9!E:F,2,FALSE),"")')
			->setCellValue('AJ' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP10!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP10!E:F,2,FALSE),"")')
			->setCellValue('AK' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP11!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP11!E:F,2,FALSE),"")')
			->setCellValue('AL' . $sheetrow, '=IF(ISNUMBER(VLOOKUP(A' . $sheetrow . ',COMP12!E:F,2,FALSE)), VLOOKUP(A' . $sheetrow . ',COMP12!E:F,2,FALSE),"")');
		$sheet->getStyle('F' . $sheetrow)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setRGB('f2f2f2');
		$sheetrow++;
		$pricerow++;		
	}
	// If products is set add the products
	if ($_POST['products']) {
		// Get list of Manufacturers Items
		$sql = "SELECT classID, partClassName FROM part_class WHERE manufactID=(" . "SELECT manufactID FROM manufacturer WHERE manufactName='" . $_POST['manufacturer'] . "');";
		$products = mysqli_query($conn, $sql);
		while ($r = mysqli_fetch_assoc($products)) {
			// Output Item Names
			addItem($r['partClassName']);
			//Get Individual Model Info
			$qry =  "SELECT p.name, pa.attriValue FROM part p LEFT JOIN part_attribute pa ON p.partID = pa.partID WHERE (pa.attriValue LIKE '"."\$%"."'
			OR pa.attriValue LIKE 'call' OR pa.attriValue LIKE 'per %' OR pa.attriValue LIKE 'request%') AND p.classID=".$r['classID']." ORDER BY p.partOrder;";
			$result = mysqli_query($conn, $qry);
		
			while ($row = mysqli_fetch_assoc($result)) {
				// Output Individual Model Info
				addModel($row['name'], floatval(preg_replace('/[\$,]/', '', $row['attriValue'])));
			}
		}
	}
	// If options is set add the options
	if ($_POST['options']) {
		$sql = "SELECT optionsID, name FROM options WHERE manufactID=("."SELECT manufactID FROM manufacturer WHERE manufactName='".$_POST['manufacturer']."');";
		$options = mysqli_query($conn, $sql);
		
		while($r = mysqli_fetch_assoc($options)) {
			addItem($r['name']);
			$qry =  "SELECT a.name name, b.optAttriValue value FROM options_part a LEFT JOIN options_part_attribute b ON a.options_partID = b.options_PartID WHERE " . 
			"(b.optAttriValue LIKE '"."\$%"."' OR b.optAttriValue LIKE 'call' OR b.optAttriValue LIKE 'per %' OR b.optAttriValue LIKE 'request%') AND a.optionsID=" . $r['optionsID'].";";
			$result = mysqli_query($conn, $qry);
    		while($row = mysqli_fetch_assoc($result)) {
				addModel($row['name'], floatval(preg_replace('/[\$,]/', '', $row['value'])));
			}
		}
	}

	// Format Data Types
	$sheet->getStyle('A5:A' . $sheetrow)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT);
	$sheet->getStyle('B5:G' . $sheetrow)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
	$sheet->getStyle('L5:W' . $sheetrow)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
	$sheet->getStyle('H5:H' . $sheetrow)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER_00);
	$sheet->getStyle('J5:K' . $sheetrow)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_PERCENTAGE_00);
	$ourprices->getStyle('A:A')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT);
	$ourprices->getStyle('B:B')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
	$msrp->getStyle('A:A')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT);
	$msrp->getStyle('B:C')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);

	// Set Row Height and Column Width
	$sheet->getColumnDimension('A')->setWidth(15);
	$sheet->getColumnDimension('B')->setWidth(12);
	$sheet->getColumnDimension('C')->setWidth(12);
	$sheet->getColumnDimension('D')->setWidth(12);
	$sheet->getColumnDimension('E')->setWidth(13);
	$sheet->getColumnDimension('F')->setWidth(12);
	$sheet->getColumnDimension('G')->setWidth(14);
	$sheet->getColumnDimension('H')->setWidth(14);
	$sheet->getColumnDimension('I')->setWidth(20);
	$sheet->getColumnDimension('J')->setWidth(8);
	$sheet->getColumnDimension('K')->setWidth(8);	
	//$sheet->getColumnDimension('J')->setAutoSize(true);
	//$sheet->getColumnDimension('K')->setAutoSize(true);
	
	// Add Macro
	$mymacro = fopen("vbaProject.bin", "r");
	$spreadsheet->setMacrosCode(fread($mymacro, filesize("vbaProject.bin")));
	$spreadsheet->setHasMacros(true);
	fclose($mymacro);
	
	// Set Default Print Area
	$sheet->getPageSetup()->setPrintArea('A1:K' . $sheetrow);

	// Set Print Margins
	$sheet->getPageMargins()->setTop(0.25)
		->setRight(0.35)
		->setLeft(0.35)
		->setBottom(0.4)
		->setFooter(0.15);
	$sheet->getPageSetup()->setHorizontalCentered(true);

	// Set Metadata
	$spreadsheet->getProperties()
    	->setCreator("Brandon Jensen")
    	->setLastModifiedBy("Brandon Jensen")
    	->setTitle($_POST['manufacturer'] . " Price List")
    	->setSubject($_POST['manufacturer'] . " Price List")
    	->setDescription($_POST['manufacturer'] . " Price List for Material Flow & Conveyor Systems, Inc.");

	// Set Footer
	$sheet->getHeaderFooter()->setOddFooter('&L&B' . $spreadsheet->getProperties()->getTitle() . '&RPage &P of &N');

	// Save And Send Spreadsheet
	$spreadsheet->setActiveSheetIndex(0);
	$writer = new Xlsx($spreadsheet);
	$writer->setPreCalculateFormulas(false); // Save the formulas to be calculated when open for speed
	$filename = $_POST['manufacturer'] . '.xlsm';
	$writer->save($filename);
  
	// Set The Content-Type:
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename=' . $filename);
	header('Content-Length: ' . filesize($filename));
	readfile($filename); // Send File
	unlink($filename); // Delete File
  
} else {
	echo "Please select items, options, or both to add to your pricing list.";
}
mysqli_close($conn); // Close all database connections, they all end here
?>