<html>
<head>

<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.1/css/bootstrap.min.css" integrity="sha384-WskhaSGFgHYWDcbwN70/dfYBj47jz9qbsMId/iRN3ewGhXQFZCSftd1LZCfmhktB" crossorigin="anonymous">

<title> Rate Your Brand </title>
<style>
.cont_style
{


text-align:left;
}
.box_style
{
	margin:5px;
}

</style>


</head>

  <body style="background-image:url('bg1.png');background-size:cover;">
<div class="container-fluid">
	
	<div class="row">
	 
	 <div class="col-sm-3">
	 <a href=".">
     <img src="Evaluate.png" class="img img-fluid"/>
	 </a>
    </div>
    <div class="col-sm-8" style="background-color:white;padding:8px;border-radius:10px;">
	<button type="button" style="float:right;" class="btn btn-info btn-sm" data-toggle="modal" data-target="#myModal">Add new Brand +</button>

	<h3>Rate Your Brands</h3>
	
	


	 <form id="fom1" name="f1" action="<?php echo $_SERVER['PHP_SELF'] ?>" method="post">

<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
// For populating the  Brands
$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$spreadsheet = $reader->load("db/brands.xlsx");
$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$worksheet = $spreadsheet->getActiveSheet();
// Get the highest row and column numbers referenced in the worksheet
$highestRow = $worksheet->getHighestRow(); // e.g. 10  // e.g. 5


// For populating the attributes
$reader1 = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader1->setReadDataOnly(TRUE);
$spreadsheet1 = $reader1->load("db/attributes.xlsx");
$writer1 = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$worksheet1 = $spreadsheet1->getActiveSheet();
// Get the highest row and column numbers referenced in the worksheet
$highestRow1 = $worksheet1->getHighestRow();

$attribute_list=array();




for ($row1 = 1; $row1 <= $highestRow1;  $row1++) {
	
	 $att_name=$worksheet1->getCell('B'.$row1)->getValue();
	 
	array_push($attribute_list,$att_name);
  
}




for ($row = 1; $row <= $highestRow;  $row++) {
  
	$id=$worksheet->getCell('A'.$row)->getValue();
			  $pname=$worksheet->getCell('B'.$row)->getValue();
			  $idname="q".$id.'[]';
			  
			  
			  
			  echo "<div class='row cont_style'>";
			  
echo "<h5 class='btn-info' style='width:100%'>  <span style='margin-left:30px;'>$pname<span> </h5>";
			  foreach($attribute_list as $value )		  
{
	
			  
	     echo "<div class='col-sm-2 box_style' style='text-align:left'><b> $value </b><select name='$idname'>
  <option value='0'>please Choose</option>
  <option value='1'>Good</option>
  <option value='2'>Very good</option>
  <option value='3'>Excellent</option>
</select></div>";
	
	
	
	 
}
			  
			  
			 
			 
		 
			  echo "</div>";

}

// a function to add a new product



function add_sheet($sheetName)
{
	
	
$readerx = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$readerx->setReadDataOnly(TRUE);
$spreadsheetx = $readerx->load("db/valuesheet.xlsx");
$writerx = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheetx, "Xlsx");
$newsheetx=$spreadsheetx->createSheet();
$newsheetx->setTitle($sheetName);
$writerx->save("db/valuesheet.xlsx");
}




function add_brand($brandName)
{
	
$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$spreadsheet = $reader->load("db/brands.xlsx");
$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$worksheet = $spreadsheet->getActiveSheet();			 
$highestRow = $worksheet->getHighestRow();

$newRow = $worksheet->getHighestRow()+1;

 $att_name=$worksheet->getCell('A'.$highestRow)->getValue();

 
 $newValue=$att_name+1;

 
$worksheet->setCellValue('A'.$newRow, $newValue);

$worksheet->setCellValue('B'.$newRow,$brandName);
$writer->save("db/brands.xlsx");

//Adding to the Excel Brands is done here and need to add a new sheet to the existing excel sheet value sheet user defined funtion


add_sheet('q'.$newValue);

 
}

//add attributeto the attirubtesxlss

function add_attribute($brandName)
{
	
$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
$reader->setReadDataOnly(TRUE);
$spreadsheet = $reader->load("db/attributes.xlsx");
$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$worksheet = $spreadsheet->getActiveSheet();			 
$highestRow = $worksheet->getHighestRow();

$newRow = $worksheet->getHighestRow()+1;

 $att_name=$worksheet->getCell('A'.$highestRow)->getValue();

 
 $newValue=$att_name+1;

 
$worksheet->setCellValue('A'.$newRow, $newValue);

$worksheet->setCellValue('B'.$newRow,$brandName);
$writer->save("db/attributes.xlsx");

//Adding to the Excel Brands is done here and need to add a new sheet to the existing excel sheet value sheet user defined funtion

 

 
}









// THis set of code for adding a new data item to the Brand list
if(isset($_POST['brandsubmit'])){
 

$bnmmm=$_POST['bname'];
add_brand($bnmmm);

}

if(isset($_POST['atsubmit'])){
 

$bnmmm=$_POST['atname'];
add_attribute($bnmmm);

}

if(isset($_POST['submit'])){
    
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
}


?>

<input type="submit" value="Submit" class="btn btn-success" name="submit"/> <input type="reset" value="Clear" class="btn btn-danger"/>
	</form>
	
	
	</div>
	 <div class="col-sm-2">
     
    </div>
	
	</div>
	
	
	</div>
	
	<!-- hhere ownards moal copy-->
	<!-- Trigger the modal with a button -->

<!-- Modal -->
<div id="myModal" class="modal fade" role="dialog">
  <div class="modal-dialog">

    <!-- Modal content-->
    <div class="modal-content">
      <div class="modal-header">
	  
	  <h4 class="modal-title">Add New Brand Name</h4>
        <button type="button" class="close" data-dismiss="modal">&times;</button>
        
      </div>
      <div class="modal-body">     
		 <form name="onlyaddform" method="post" action="<?php echo $_SERVER['PHP_SELF']; ?>"  style="width:100%">
		 
		<input  type="text" name="bname" class="form-control" style="width:100%;" placeholder="Enter your new BrandName" required/>
		
		 <p>
		<input type="submit" name="brandsubmit" value="Add New Brand" class="btn btn-success btn-sm" style="float:right; margin:5px;"/>
		
		</p>
		</form>
		
			<form name="onlyaddform2" method="post" action="<?php echo $_SERVER['PHP_SELF']; ?>"  style="width:100%">
		 
		<input  type="text" name="atname" class="form-control" style="width:100%;" placeholder="Enter a New Attribute" required/>
		
		 <p>
		<input type="submit" name="atsubmit" value="Add New Attribute" class="btn btn-success btn-sm" style="float:right; margin:5px;"/>
		
		</p>
		</form>
		
		
		 
      </div>
      <div class="modal-footer"> 
      </div>
    </div>

  </div>
</div>
	
<script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.3/umd/popper.min.js" integrity="sha384-ZMP7rVo3mIykV+2+9J3UJ46jBk0WLaUAdn689aCwoqbBJiSnjAK/l8WvCWPIPm49" crossorigin="anonymous"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.1.1/js/bootstrap.min.js" integrity="sha384-smHYKdLADwkXOn1EmN1qk/HfnUcbVRZyYmZ4qpPea6sjB/pTJ0euyQp0Mk8ck+5T" crossorigin="anonymous"></script>
	
	
</body>
</html>