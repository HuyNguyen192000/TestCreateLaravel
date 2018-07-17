<?php

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/excel-for-GK', function () {
	$inputFileType = 'Xls';
	// $inputFileName = public_path() . '/files/DMSACHGK_fix8.Xls';
	$inputFileName = public_path() . '/files/DMSACH_fix8.Xls';
	/**  Create a new Reader of the type defined in $inputFileType  **/
	$objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($inputFileName);
	$objReader->setReadDataOnly(true);
	$objPHPExcel = $objReader->load($inputFileName);

	$sheet = $objPHPExcel->getActiveSheet();
	$highestRow = $sheet->getHighestRow();
	$highestColumn = $sheet->getHighestColumn();
	$rowData = [];
	for ( $row = 2 ; $row <= $highestRow; $row++) {
		$rowExcel = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);
		$rowData[$row-2]['Mã đoạn'] = $rowExcel[0][0];
		$rowData[$row-2]['Mã ĐKCB'] = $rowExcel[0][1];
		$rowData[$row-2]['Tên Sách'] = $rowExcel[0][2];
		$rowData[$row-2]['Nhan đề song song'] = $rowExcel[0][3];
		$rowData[$row-2]['Kho sách
(Chỉ nhập ký hiệu kho ví dụ: GK, TN…)'] = $rowExcel[0][4];
		$rowData[$row-2]['Mô tả - phụ đề'] = $rowExcel[0][5];
		$rowData[$row-2]['Tập'] = $rowExcel[0][6];
		$rowData[$row-2]['Tên tập'] = $rowExcel[0][7];
		$rowData[$row-2]['Tên tác giả'] = $rowExcel[0][8];
		$rowData[$row-2]['Đồng tác giả'] = $rowExcel[0][9];
		$rowData[$row-2]['Dịch giả'] = $rowExcel[0][10];
		$rowData[$row-2]['Nơi XB'] = $rowExcel[0][1];
		$rowData[$row-2]['NXB'] = $rowExcel[0][12];
		$rowData[$row-2]['Năm XB
(Chỉ nhập số)'] = $rowExcel[0][13];
		$rowData[$row-2]['Lần XB
(Chỉ nhập số)'] = $rowExcel[0][14];
		$rowData[$row-2]['Giá bìa
(Chỉ nhập số)'] = $rowExcel[0][15];
		$rowData[$row-2]['Mã phân loại
(Chỉ nhập ký hiệu phân loại)'] = $rowExcel[0][16];
		$rowData[$row-2]['Số trang
(Chỉ nhập số)'] = $rowExcel[0][17];
		$rowData[$row-2]['Khổ cỡ'] = $rowExcel[0][18];
		$rowData[$row-2]['Số sổ ĐKTQ'] = $rowExcel[0][19];
	}
	//
	$arrayBooksName = [];
	foreach ($rowData as $key => $value) {
		$arrayBooksName[] = $value['Tên Sách'];
	}
	$countBooksName = array_count_values($arrayBooksName);

	$newRowData = [];
	foreach ($rowData as $key => $value) {
		$newRowData[$value['Tên Sách']] = $value;
	}
	// xao mang
	$keys = array_keys($newRowData); 
	shuffle($keys); 
	$randomBook = []; 
	foreach ($keys as $key) {
	 	$randomBook[$key] = $newRowData[$key]; 
	}
	// dd($randomBook);
	$arr = [];
	$n = 0;
	foreach ($randomBook as $key => $value) {
		foreach ($countBooksName as $key_count => $value_count) {
			if ( $key == $key_count ) {
				$arr[$n] = $value;
				$arr[$n]['sl'] = $value_count;
			}
		}
		$n++;
	}
	// delete random 
	$numberRandom = rand(1, (count($arr) / 2));
	$deleleKeys = array_rand($arr, $numberRandom);
	foreach ($deleleKeys as $value) {
		unset($arr[$value]);
	}

	// get array output 
	$y = 0;
	$newArr = [];
	foreach ($arr as $key => $value) {
		$n = 0 ;
		for ($i=1; $i <= $value['sl']; $i++) {
			$newArr[$y] = $value;
			$newArr[$y]['Mã ĐKCB'] = 'GK' . str_pad(( $y + 1 ), 5, 0 , STR_PAD_LEFT);
			//
			$begin = $y - $n + 1;
			$end = $y + $value['sl'] - $n ;
			if ($value['sl'] == 1) {
				$newArr[$y]['Mã đoạn'] = 'GK' . str_pad(( $begin ), 5, 0 , STR_PAD_LEFT);
			} else {
				$newArr[$y]['Mã đoạn'] = 'GK' . str_pad(( $begin ), 5, 0 , STR_PAD_LEFT) . '-' .
						'GK' . str_pad(( $end), 5, 0 , STR_PAD_LEFT);
			}
			
			$n++ ;
			if ($n == $value['sl']) {
				$n = 0;
			}
			$y++;
		}
	}
	// dd($newArr);
	$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileName);
	// $spreadsheet->getActiveSheet()->removeRow(2, 589);
	$spreadsheet->getActiveSheet()->removeRow(2, 2065);
	$sheet = $spreadsheet->getActiveSheet();

	foreach ($newArr as $key => $value) {
		$sheet->setCellValue('A' . ($key + 2) , $value['Mã đoạn']);
		$sheet->setCellValue('B' . ($key + 2) , $value['Mã ĐKCB']);
		$sheet->setCellValue('C' . ($key + 2) , $value['Tên Sách']);
		$sheet->setCellValue('D' . ($key + 2) , $value['Nhan đề song song']);
		$sheet->setCellValue('E' . ($key + 2) , $value['Kho sách
(Chỉ nhập ký hiệu kho ví dụ: GK, TN…)']);
		$sheet->setCellValue('F' . ($key + 2) , $value['Mô tả - phụ đề']);
		$sheet->setCellValue('G' . ($key + 2) , $value['Tập']);
		$sheet->setCellValue('H' . ($key + 2) , $value['Tên tập']);
		$sheet->setCellValue('I' . ($key + 2) , $value['Tên tác giả']);
		$sheet->setCellValue('J' . ($key + 2) , $value['Đồng tác giả']);
		$sheet->setCellValue('K' . ($key + 2) , $value['Dịch giả']);
		$sheet->setCellValue('L' . ($key + 2) , $value['Nơi XB']);
		$sheet->setCellValue('M' . ($key + 2) , $value['NXB']);
		$sheet->setCellValue('N' . ($key + 2) , $value['Năm XB
(Chỉ nhập số)']);
		$sheet->setCellValue('O' . ($key + 2) , $value['Lần XB
(Chỉ nhập số)']);
		$sheet->setCellValue('P' . ($key + 2) , $value['Giá bìa
(Chỉ nhập số)']);
		$sheet->setCellValue('Q' . ($key + 2) , $value['Mã phân loại
(Chỉ nhập ký hiệu phân loại)']);
		$sheet->setCellValue('R' . ($key + 2) , $value['Số trang
(Chỉ nhập số)']);
		$sheet->setCellValue('S' . ($key + 2) , $value['Khổ cỡ']);
		$sheet->setCellValue('T' . ($key + 2) , $value['Số sổ ĐKTQ']);
	}
	// download 
	$outputFileName = public_path() . '/output/GK_' . date('YmdHis') . '.xls';
	$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xls');
	$writer->save($outputFileName);
	return response()->download($outputFileName);

    
});

Route::get('/', function(){
	return view('welcome');
});

Route::get('/test', function(){
	// try {
	// 	\DB::connection()->getPdo();
	// } catch (Exception $e) {
	// 	die("Could not connect to the database.  Please check your configuration.");
	// }
	\DB::connection()->getPdo();
	dd(\DB::table('categories')->get());
	dd(1);
});

Route::get('/excel', function(){
	dd(1);
	$inputFileName = public_path() . '/files/DMSACH_fix8.Xls';
	/**  Create a new Reader of the type defined in $inputFileType  **/
	$objReader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($inputFileName);
	$objReader->setReadDataOnly(true);
	$objPHPExcel = $objReader->load($inputFileName);
});
