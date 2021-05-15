<!--Deklarasi Script PHP-->
<?php
	//membuka koneksi ke database
	include "koneksi.php";
	
	//memanggil file library spreadsheet
	require 'vendor/autoload.php';
	use PhpOffice\PhpSpreadsheet\Spreadsheet;
	use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

	//Deklarasi variabel dengan class spreadsheet
	$spreadsheet=new spreadsheet();
	//Deklarasi var sheet sebagai activesheet di file excel
	$sheet=$spreadsheet->getActiveSheet();
	
	//Membuat judul pada setiap kolom
	$sheet->setCellValue('A1','No');
	$sheet->setCellValue('B1','Nama');
	$sheet->setCellValue('C1','Kelas');
	$sheet->setCellValue('D1','Alamat');

	//mengambil data pada database dan disimpan di var query
	$query=mysqli_query($conn,"select * from tb_siswa");
	//menyimpan nomor awal cell dan ditampilkan mulai baris 2
	$i=2;
	$no=1;
	//extract hasil query dan data disimpan di var row
	while($row=mysqli_fetch_array($query)){
		$sheet->setCellValue('A'.$i,$no++);
		$sheet->setCellValue('B'.$i,$row['nama']);
		$sheet->setCellValue('C'.$i,$row['kelas']);
		$sheet->setCellValue('D'.$i,$row['alamat']);
		$i++;
	}

	//style border untuk cell
	$styleArray=[
				'borders'=>[
					'allBorders'=>[
						'borderStyle'=>PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
					],
				],
	];

	//memunculkan file excel
	$i=$i-1;
	$sheet->getStyle('A1:D'.$i)->applyFromArray($styleArray);
	//Merender menjadi file xlsx
	$writer=new Xlsx($spreadsheet);
	//Ekspor file excel
	$writer->save('Report Data Siswa.xlsx');
?>