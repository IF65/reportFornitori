<?php
	//@ini_set('memory_limit','8192M');

	require(realpath(__DIR__ . '/..').'/vendor/autoload.php');
    require(realpath(__DIR__ . '/..').'/src/Database/autoload.php');
	
    use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
    use PhpOffice\PhpSpreadsheet\Cell\DataType;
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    use PhpOffice\PhpSpreadsheet\Style\Style;
    use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
    use PhpOffice\PhpSpreadsheet\Style\Alignment;
    use PhpOffice\PhpSpreadsheet\Style\Fill;
    use PhpOffice\PhpSpreadsheet\Style\Border;
	use PhpOffice\PhpSpreadsheet\Shared\Date;

	use Database\Database;

	$timeZone = new \DateTimeZone('Europe/Rome');
	
	$db = new Database($sqlDetails);
	
	$response = $db->huawei->caricaDati(['dallaData' => '2018-08-01', 'allaData' => '2018-09-01']);
	
	
	// creo il file excel
	$style = new Style();
	
	// creazione del workbook
    $workBook = new Spreadsheet();
    $workBook->getDefaultStyle()->getFont()->setName('Arial');
    $workBook->getDefaultStyle()->getFont()->setSize(12);
    $workBook->getProperties()
        ->setCreator("Supermedia S.p.A. (Gruppo Italmark)")
        ->setLastModifiedBy("Supermedia S.p.A.")
        ->setTitle("Report Vendite Huawei")
        ->setSubject("Report Vendite Huawei")
        ->setDescription("Report Vendite Huawei")
        ->setKeywords("office 2007 openxml php")
        ->setCategory("SM Docs");
		
	$sheet = $workBook->setActiveSheetIndex(0); // la numerazione dei worksheet parte da 0
    $sheet->setTitle('Report vendite Huawei');
	
	// testata
	$sheet->getCell('A1')->setValueExplicit('anno',DataType::TYPE_STRING);
	$sheet->getCell('B1')->setValueExplicit('settimana',DataType::TYPE_STRING);
	$sheet->getCell('C1')->setValueExplicit('data',DataType::TYPE_STRING);
	$sheet->getCell('D1')->setValueExplicit('cliente',DataType::TYPE_STRING);
	$sheet->getCell('E1')->setValueExplicit('negozio',DataType::TYPE_STRING);
	$sheet->getCell('F1')->setValueExplicit('negozioDescrizione',DataType::TYPE_STRING);
	$sheet->getCell('G1')->setValueExplicit('barcode',DataType::TYPE_STRING);
	$sheet->getCell('H1')->setValueExplicit('codice',DataType::TYPE_STRING);
	$sheet->getCell('I1')->setValueExplicit('descrizione',DataType::TYPE_STRING);
	$sheet->getCell('J1')->setValueExplicit('quantita',DataType::TYPE_STRING);
	$sheet->getCell('K1')->setValueExplicit('stock',DataType::TYPE_STRING);
	$sheet->getCell('L1')->setValueExplicit('importoTotale',DataType::TYPE_STRING);
		
	foreach ($response['data'] as $rowIndex => $row) {
		$r = sprintf('%d', $rowIndex + 2);
		$sheet->getCell('A'.$r)->setValueExplicit($row['anno'],DataType::TYPE_NUMERIC);
		$sheet->getCell('B'.$r)->setValueExplicit($row['settimana'],DataType::TYPE_NUMERIC);
		$sheet->getCell('C'.$r)->setValueExplicit($row['data'],DataType::TYPE_STRING);
		$sheet->getCell('D'.$r)->setValueExplicit($row['cliente'],DataType::TYPE_STRING);
		$sheet->getCell('E'.$r)->setValueExplicit($row['negozio'],DataType::TYPE_STRING);
		$sheet->getCell('F'.$r)->setValueExplicit($row['negozio_descrizione'],DataType::TYPE_STRING);
		$sheet->getCell('G'.$r)->setValueExplicit($row['barcode'],DataType::TYPE_STRING);
		$sheet->getCell('H'.$r)->setValueExplicit($row['codice'],DataType::TYPE_STRING);
		$sheet->getCell('I'.$r)->setValueExplicit($row['descrizione'],DataType::TYPE_STRING);
		$sheet->getCell('J'.$r)->setValueExplicit($row['quantita'],DataType::TYPE_NUMERIC);
		$sheet->getCell('K'.$r)->setValueExplicit($row['stock'],DataType::TYPE_NUMERIC);
		$sheet->getCell('L'.$r)->setValueExplicit($row['importoTotale'],DataType::TYPE_NUMERIC);
	}
	
	$writer = new Xlsx($workBook);
    $writer->save('/Users/if65/Desktop/huawei.xlsx');