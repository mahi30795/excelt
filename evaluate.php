<?php
	 
	  require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
	
		 
	
// For populating the attributes
$reader1 = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader1->setReadDataOnly(TRUE);
$spreadsheet1 = $reader1->load("db/brands.xlsx");
$writer1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet1, "Xlsx");
$worksheet1 = $spreadsheet1->getActiveSheet();
// Get the highest row and column numbers referenced in the worksheet
$highestRow1 = $worksheet1->getHighestRow();

$attribute_list=array();




for ($row1 = 1; $row1 <= $highestRow1;  $row1++) {
	

			  $ppid=$worksheet1->getCell('A'.$row1)->getValue();
			  $postIndex='q'.$ppid;
			  $name_parsing[$row1-1]=$_POST[$postIndex];
			  
			  
  
}

 
	// test code
	
$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);

	
	
	
	  
			  		  
					  
			 	
				for ($i = 0; $i <$highestRow1; $i++) {
					
					// Index needed to be used in multiple place
					$sheetIndex="q".$i;
				// echo "The value of last rw- $lastRow";
				
				
				$spreadsheet = $reader->load("db/valuesheet.xlsx");
				$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
				 
				
				//new code
				 
				
			 	$sheet=$spreadsheet->setActiveSheetIndexByName($sheetIndex);
				
			 	$row = $sheet->getHighestRow()+1;
				
				
				 
				
			  foreach( $name_parsing[$i] as $key => $n ) {
				  
				//print "<br>The Rating of Some item  for Key-> $key is  value-> $n:</br>";
				
				$sheet->setCellValueByColumnAndRow($key+1, $row,$n );
				
				
					}
			    $writer->save("db/valuesheet.xlsx");
					   
				}

				
			 
 
?>



<!doctype html>
<html lang="en">
  <head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.1/css/bootstrap.min.css" integrity="sha384-WskhaSGFgHYWDcbwN70/dfYBj47jz9qbsMId/iRN3ewGhXQFZCSftd1LZCfmhktB" crossorigin="anonymous">

    <title>Thank You For your Ratings</title>
	<style>
	
	#fom1 p
	{
		margin-top:30px;
		margin-bottom:30px;
		padding:6px;
		border-bottom:1px solid black;
	}
	
	</style>
	
  </head>
  <body style="background-image:url('bg1.png');background-size:cover;">
  
  
  
  
    
	<div class="container-fluid">
	 
	<div class="row">
	 
	<div class="col-sm-12 text-center" style=""> 
	 <a href=".">
     <img src="Evaluate.png" class="img img-fluid"/>
	 </a>
	
	
	<h1 style="color:white;text-shadow:2px 2px #3c2c2c;">
	 Thank You 
	 </br>
	Your Evaluation is Submitted Successfully
	</br>
	
	<a href="./"> (click here) to Go back</a>	
	</h1>
	</div>
 
	  
	</div>
	
	</div>
	
	</body>
	
	
	</html>
	







