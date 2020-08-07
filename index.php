<?php
	include 'vendor/autoload.php';

	use PhpOffice\PhpSpreadsheet\Spreadsheet;
	use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
	$conn = mysqli_connect("localhost", "phpmyadmin", "yayati23", "market");
	$query = "select * from driver_settlement where driver_phone='9876543210' and date_of_settlement between '2020-07-30' and '2020-08-02' order by settlement_id desc";
	$result = mysqli_query($conn, $query);
	
	if(isset($_POST['export']))
	{
		$file = new Spreadsheet();
		$active_sheet = $file->getActiveSheet();
		$active_sheet->setCellValue('A1', 'Id');
		$active_sheet->setCellValue('B1', 'Order ID');
		$active_sheet->setCellValue('C1', 'Amount To Be Received');
		$active_sheet->setCellValue('D1', 'Status');
		$active_sheet->setCellValue('E1', 'Date of Settlement');
		$active_sheet->setCellValue('F1', 'Account Number');

		$total_count = 2; 
		$count = 1;
		while($row = mysqli_fetch_array($result))
		{
			$active_sheet->getStyle('A:B')->getAlignment()->setHorizontal('center');
			$active_sheet->getStyle('C:D')->getAlignment()->setHorizontal('center');
			$active_sheet->getStyle('E:F')->getAlignment()->setHorizontal('center');
			
			$active_sheet->setCellValue('A'.$total_count, $count);
			$active_sheet->setCellValue('B'.$total_count, $row['order_id']);
			$active_sheet->setCellValue('C'.$total_count, $row['amount']);
			$active_sheet->setCellValue('D'.$total_count, $row['status']);
			$active_sheet->setCellValue('E'.$total_count, $row['date_of_settlement']);
			$active_sheet->setCellValue('F'.$total_count, $row['account_number']);

			//$active_sheet->getColumnDimension('A')->setAutoSize('false');
			$active_sheet->getColumnDimension('A')->setWidth(5);
			$active_sheet->getColumnDimension('B')->setAutoSize('false');
			// $active_sheet->getColumnDimension('B')->setWidth(50);
			//$active_sheet->getColumnDimension('C')->setAutoSize('false');
			$active_sheet->getColumnDimension('C')->setWidth(21);
			// $active_sheet->getColumnDimension('D')->setAutoSize('false');
			$active_sheet->getColumnDimension('D')->setWidth(21);
			// $active_sheet->getColumnDimension('E')->setAutoSize('false');
			$active_sheet->getColumnDimension('E')->setWidth(17);
			// $active_sheet->getColumnDimension('F')->setAutoSize('false');
			$active_sheet->getColumnDimension('F')->setWidth(16);
			
			$total_count += 1;
			$count += 1;
		}
		$filename = time().'.'.strtolower('Xlsx');
		

		$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($file, 'Xlsx');
		ob_end_clean();
		header('Content-Type: application/x-www-form-urlencoded');
		header('Content-Transfer-Encoding: Binary');
		header("Content-disposition: attachment; filename = \"".$filename. "\"");
		header('Cache-Control: max-age=0');
		$writer->save('php://output');
		// $writer->save($filename);
		// readfile($filename);
		// unlink($filename);
		exit();
	}
?>
<html>
	<head>
	</head>
	<body>
		<form method="post">
			<select name="file_type">
				<option value="Xlsx">Xlsx</option>
				<option value="Xls">Xls</option>
				<option value="Csv">Csv</option>
			</select>
			<input type="submit" name="export" class="" value="Export">
		</form>
		<?php
		$count = 1;
		$output ="";	
		if(mysqli_num_rows($result) > 0)
        {
            $output .= '
                <table border="1">
                    <tr>
                        <th>Id</th>
                        <th>Order ID</th>
                        <th>Amount To Be Received</th>
                        <th>Status</th>
                        <th>Date of Settlement</th>
                        <th>Account Number</th>
                    </tr>
            ';
            while($row = mysqli_fetch_array($result))
            {
                $output .= '
                    <tr>
                        <td>'. $count++.'</td>
                        <td>'. $row["order_id"].'</td>
                        <td>'. $row["amount"].'</td>
                        <td>'. $row["status"].'</td>
                        <td>'. $row["date_of_settlement"].'</td>
                        <td>'. $row["account_number"].'</td>
                    </tr>
                ';
            }
            $output .= '</table>';
		}
		echo $output;
		?>
	</body>
</html>
	
