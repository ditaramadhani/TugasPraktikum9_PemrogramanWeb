<!--Deklarasi Script PHP-->
<?php
	//memanggil file library 
	require 'vendor/autoload.php';
	use PhpOffice\PhpSpreadsheet\Spreadsheet;
	use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

	//membuat konstruktor
	$spreadsheet = new Spreadsheet();
	$sheet = $spreadsheet->getActiveSheet();
	//Membuat teks Hello World pada cell A1
	$sheet->setCellValue('A1','Hello World!');

	//Membuat dan memberi nama file excel dengan format xlsx
	$writer = new Xlsx($spreadsheet);
	$writer -> save('hello world.xlsx');
?>
