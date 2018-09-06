<?php
	//@ini_set('memory_limit','8192M');

	require '../vendor/autoload.php';
	// leggo i dati da un file
    //$request = file_get_contents('../examples/ordini.json');
    $request = file_get_contents('php://input');
    $data = json_decode($request, true);

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

	// verifico l'esistenza della cartella temp e se serve la creo
	// con mask 777.
	if (! file_exists ( '../temp' )) {
		$oldMask = umask(0);
		mkdir('../temp', 0777);
		umask($oldMask);
	}

    $style = new Style();

    // leggo i parametri contenuti nel file
    $nomeFile = $data['nomeFile'];
    $file = '../temp/'.$nomeFile.'.xlsx';

    $ordini = $data['ordini'];
    $ordinamento = array();
    foreach ($ordini as $key => $row) {
        $ordinamento[$key] = $row['numero'];
    }
    array_multisort($ordinamento, SORT_ASC, $ordini);

    // creazione del workbook
    $workBook = new Spreadsheet();
    $workBook->getDefaultStyle()->getFont()->setName('Arial');
    $workBook->getDefaultStyle()->getFont()->setSize(12);
    $workBook->getProperties()
        ->setCreator("IF65 S.p.A. (Gruppo Italmark)")
        ->setLastModifiedBy("IF65 S.p.A.")
        ->setTitle("Ordine Acquisto")
        ->setSubject("Ordine Acquisto")
        ->setDescription("Esportazione Ordine di Acquisto")
        ->setKeywords("office 2007 openxml php")
        ->setCategory("SM Docs");

    // creazione degli Sheet (uno per ogni ordine)
    $sheetNumber = 0;
    foreach ($ordini as $ordine) {
        $sheetNumber++;
        if ($workBook->getSheetCount() < $sheetNumber) {
            $workBook->createSheet();
        }
        $sheet = $workBook->setActiveSheetIndex($sheetNumber-1); // la numerazione dei worksheet parte da 0
        $sheet->setTitle(preg_replace('/\//','_',$ordine['numero']));

		$timeZone = new DateTimeZone('Europe/Rome');

		$dataOrdine = new \DateTime($ordine['data']);
		$dataConsegna= new \DateTime($ordine['dataConsegna']);
		$dataConsegnaMinima= new \DateTime($ordine['dataConsegnaMinima']);
		$dataConsegnaMassima= new \DateTime($ordine['dataConsegnaMassima']);

		$filiali = $ordine['sedi'];
		$ordinamento = array();
		foreach ($filiali as $key => $row) {
			$ordinamento[$key] = $row['ordinamento'];
		}
		array_multisort($ordinamento, SORT_ASC, $filiali);
		$countFiliali = count($filiali);


        // riquadro di testata
        // --------------------------------------------------------------------------------
        $sheet->setCellValue('A1', strtoupper('fornitore'));
        $sheet->setCellValue('B1', $ordine['fornitore']);
        $sheet->mergeCells('B1:C1');
        $sheet->setCellValue('D1',strtoupper('forma di pagamento'));
        $sheet->setCellValue('E1', $ordine['pagamento']);
        $sheet->setCellValue('A2', strtoupper('numero ordine'));
        $sheet->setCellValue('B2', $ordine['numero']);
        $sheet->mergeCells('B2:C2');
        $sheet->setCellValue('D2', strtoupper('sconto cassa %'));
        $sheet->setCellValue('E2', $ordine['scontoCassa']);
        $sheet->setCellValue('A3',strtoupper('data ordine'));
        $sheet->setCellValue('B3', Date::PHPToExcel($dataOrdine->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getStyle('B3')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
		$sheet->mergeCells('B3:C3');
        $sheet->setCellValue('D3', strtoupper('spese di trasporto'));
        $sheet->setCellValue('E3', $ordine['speseTrasporto']);
        $sheet->setCellValue('A4', strtoupper('data consegna prevista'));
		$sheet->setCellValue('B4', Date::PHPToExcel($dataConsegna->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getStyle('B4')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
        $sheet->mergeCells('B4:C4');
        $sheet->setCellValue('D4', strtoupper('spese di trasporto %'));
        $sheet->setCellValue('E4', $ordine['speseTrasportoPerc']);
        $sheet->setCellValue('A5', strtoupper('data consegna minima'));
		$sheet->setCellValue('B5', Date::PHPToExcel($dataConsegnaMinima->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getStyle('B5')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
        $sheet->mergeCells('B5:C5');
        $sheet->setCellValue('D5', strtoupper('margine totale'));
        $sheet->setCellValue('E5', 0); //piu' avanti inserita la formuladi calcolo
        $sheet->setCellValue('A6', strtoupper('data consegna massima'));
        $sheet->setCellValue('B6', Date::PHPToExcel($dataConsegnaMassima->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getStyle('B6')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
        $sheet->mergeCells('B6:C6');
        $sheet->setCellValue('D6', strtoupper('margine %'));
        $sheet->setCellValue('E6', 0); //piu' avanti inserita la formuladi calcolo
        $sheet->setCellValue('A7', strtoupper('buyer'));
        $sheet->setCellValue('B7', $ordine['buyerCodice'].' - '.$ordine['buyerDescrizione']);
        $sheet->mergeCells('B7:C7');
        $sheet->setCellValue('D7', strtoupper('totale ordine'));
        $sheet->setCellValue('E7', 0); //piu' avanti inserita la formuladi calcolo

        // testata colonne
        // --------------------------------------------------------------------------------
        $sheet->setCellValue('A9', strtoupper('cod.art. fornitore'));
        $sheet->mergeCells('A9:A10');
        $sheet->setCellValue('B9', strtoupper('ean'));
        $sheet->mergeCells('B9:B10');
        $sheet->setCellValue('C9', strtoupper('cod. art.'));
        $sheet->mergeCells('C9:C10');
        $sheet->setCellValue('D9', strtoupper('descrizione'));
        $sheet->mergeCells('D9:D10');
        $sheet->setCellValue('E9', strtoupper('marca'));
        $sheet->mergeCells('E9:E10');
        $sheet->setCellValue('F9', strtoupper('modello'));
        $sheet->mergeCells('F9:F10');
        $sheet->setCellValue('G9', strtoupper('fam.'));
        $sheet->mergeCells('G9:G10');
        $sheet->setCellValue('H9', strtoupper('s.fam.'));
        $sheet->mergeCells('H9:H10');
        $sheet->setCellValue('I9', strtoupper('iva'));
        $sheet->mergeCells('I9:J9');
        $sheet->setCellValue('I10', '%');
        $sheet->setCellValue('J10', 'T');
        $sheet->setCellValue('K9', strtoupper('tg.'));
        $sheet->mergeCells('K9:K10');
        $sheet->setCellValue('L9', strtoupper('costo'));
        $sheet->mergeCells('L9:L10');
        $sheet->setCellValue('M9', strtoupper('sconti'));
        $sheet->mergeCells('M9:R9');
        $sheet->setCellValue('M10', 'A');
        $sheet->setCellValue('N10', 'B');
        $sheet->setCellValue('O10', 'C');
        $sheet->setCellValue('P10', 'D');
        $sheet->setCellValue('Q10', 'EXT.');
        $sheet->setCellValue('R10', strtoupper('imp.'));
        $sheet->setCellValue('S9', strtoupper('costo finito'));
        $sheet->mergeCells('S9:S10');
        $sheet->setCellValue('T9', strtoupper('prezzo vendita'));
        $sheet->mergeCells('T9:T10');
        $sheet->setCellValue('U9', strtoupper('marg.'));
        $sheet->mergeCells('U9:U10');
        $sheet->setCellValue('V9', strtoupper('marg.%'));
        $sheet->mergeCells('V9:V10');
        $sheet->setCellValue('W9', strtoupper('marg. totale'));
        $sheet->mergeCells('W9:W10');
        $sheet->setCellValue('X9', strtoupper('costo totale'));
        $sheet->mergeCells('X9:X10');

        $sheet->mergeCells('A8:X8');
        $sheet->mergeCells('F1:X7');

        // testata quantita
        // --------------------------------------------------------------------------------
        $col = 'Y';
        $colQuantitaTotale = $col;
        $sheet->setCellValue($col.'3', strtoupper('totale pezzi'))
            ->getStyle($col.'3')
            ->getAlignment()
            ->setTextRotation(90)
            ->setHorizontal('center')
            ->setVertical('bottom');
        $sheet->mergeCells($col.'3:'.$col.'10');

        $col++;
        $colQIndex = array('FIRST' => $col);
        $codiciFiliali = array_keys($filiali);
        for ($i=0;$i<count($codiciFiliali);$i++) {
            $sheet->setCellValue($col.'3', strtoupper($filiali[$codiciFiliali[$i]]['descrizione']))
                ->getStyle($col.'3')
                ->getAlignment()
                ->setTextRotation(90)
                ->setHorizontal('center')
                ->setVertical('bottom');

            $sheet->mergeCells($col.'3:'.$col.'10');

            $colQIndex[$codiciFiliali[$i]] = $col;
            $colQIndex['LAST'] = $col;

            $col++;
        }
        $sheet->setCellValue($colQuantitaTotale.'1', strtoupper('QUANTITA\' IN ORDINE'));
        $sheet->mergeCells(sprintf("%s%s%s%s",$colQuantitaTotale,'1:',$colQIndex['LAST'],2));

        // testata sconto merce
        // --------------------------------------------------------------------------------
        $colScontoMerceTotale = $col;
        $sheet->setCellValue($col.'3', strtoupper('totale sconto merce'))
            ->getStyle($col.'3')
            ->getAlignment()
            ->setTextRotation(90)
            ->setHorizontal('center')
            ->setVertical('bottom');
        $sheet->mergeCells($col.'3:'.$col.'10');

        $col++;
        $colSCIndex = array('FIRST' => $col);
        $codiciFiliali = array_keys($filiali);
        for ($i=0;$i<count($codiciFiliali);$i++) {
            $sheet->setCellValue($col.'3', strtoupper($filiali[$codiciFiliali[$i]]['descrizione']))
                ->getStyle($col.'3')
                ->getAlignment()
                ->setTextRotation(90)
                ->setHorizontal('center')
                ->setVertical('bottom');

            $sheet->mergeCells($col.'3:'.$col.'10');

            $colSCIndex[$codiciFiliali[$i]] = $col;
            $colSCIndex['LAST'] = $col;

            $col++;
        }

        $sheet->setCellValue($colScontoMerceTotale.'1', strtoupper('QUANTITA\' IN SCONTO MERCE'));
        $sheet->mergeCells(sprintf("%s%s%s%s",$colScontoMerceTotale,'1:',$colSCIndex['LAST'],2));

        // scrittura righe
        // --------------------------------------------------------------------------------
        $primaRigaDati = 11; // attenzione le righe in Excel partono da 1

        $righe = $ordine['righe'];
    	$ordinamento = array();
    	foreach ($righe as $key => $row) {
        	$ordinamento[$key] = $row['codice'];
    	}
    	array_multisort($ordinamento, SORT_ASC, $righe);

        for ($i = 0; $i < count($righe); $i++) {
            $R = ($i+$primaRigaDati);
            $QFirst = $colQIndex['FIRST'];
            $QLast = $colQIndex['LAST'];
            $SCFirst = $colSCIndex['FIRST'];
            $SCLast = $colSCIndex['LAST'];

            // formule
            $costoFinito ="=IF(BM$R+Y$R>0,ROUND((L$R*(100-M$R)/100*(100-N$R)/100*(100-O$R)/100*(100-P$R)/100*(100-Q$R)/100-R$R)*Y$R/(BM$R+Y$R),2),ROUND((L$R*(100-M$R)/100*(100-N$R)/100*(100-O$R)/100*(100-P$R)/100*(100-Q$R)/100-R$R),2))";
            $margine = "=ROUND(T$R*(100/(100+I$R))-S$R,2)";
            $marginePercentuale = "=IF(T$R<>0,ROUND(U$R/(T$R*(100/(100+I$R)))*100,2),0)";
            $margineTotale = "=ROUND(U$R*(Y$R+BM$R),2)";
            $costoTotale = "=ROUND(S$R*(Y$R+BM$R),2)";
            $quantitaTotale = "=SUM($QFirst$R:$QLast$R)";
            $scontoMerceTotale = "=SUM($SCFirst$R:$SCLast$R)";

            // righe
            $sheet->getCell('A'.$R)->setValueExplicit($righe[$i]['codiceArticoloFornitore'],DataType::TYPE_STRING);
            $barcode = $righe[$i]['barcode'];
            if (count($barcode)) {
                $sheet->getCell('B'.$R)->setValueExplicit($barcode[0],DataType::TYPE_STRING);
            }
            $sheet->getCell('C'.$R)->setValueExplicit($righe[$i]['codice'],DataType::TYPE_STRING);
            $sheet->getCell('D'.$R)->setValueExplicit($righe[$i]['descrizione'],DataType::TYPE_STRING);
            $sheet->getCell('E'.$R)->setValueExplicit($righe[$i]['marca'],DataType::TYPE_STRING);
            $sheet->getCell('F'.$R)->setValueExplicit($righe[$i]['modello'],DataType::TYPE_STRING);
            $sheet->getCell('G'.$R)->setValueExplicit($righe[$i]['famiglia'],DataType::TYPE_STRING);
            $sheet->getCell('H'.$R)->setValueExplicit($righe[$i]['sottoFamiglia'],DataType::TYPE_STRING);
            $sheet->getCell('I'.$R)->setValueExplicit($righe[$i]['iva'],DataType::TYPE_NUMERIC);
            $sheet->getCell('J'.$R)->setValueExplicit($righe[$i]['tipoIva'],DataType::TYPE_NUMERIC);
            $sheet->getCell('K'.$R)->setValueExplicit($righe[$i]['taglia'],DataType::TYPE_NUMERIC);
            $sheet->getCell('L'.$R)->setValueExplicit($righe[$i]['listino'],DataType::TYPE_NUMERIC);
            $sheet->getCell('M'.$R)->setValueExplicit($righe[$i]['scontoA'],DataType::TYPE_NUMERIC);
            $sheet->getCell('N'.$R)->setValueExplicit($righe[$i]['scontoB'],DataType::TYPE_NUMERIC);
            $sheet->getCell('O'.$R)->setValueExplicit($righe[$i]['scontoC'],DataType::TYPE_NUMERIC);
            $sheet->getCell('P'.$R)->setValueExplicit($righe[$i]['scontoD'],DataType::TYPE_NUMERIC);
            $sheet->getCell('Q'.$R)->setValueExplicit($righe[$i]['scontoExtra'],DataType::TYPE_NUMERIC);
            $sheet->getCell('R'.$R)->setValueExplicit($righe[$i]['scontoImporto'],DataType::TYPE_NUMERIC);
            $sheet->getCell('S'.$R)->setValueExplicit($costoFinito,DataType::TYPE_FORMULA);
            $sheet->getCell('T'.$R)->setValueExplicit($righe[$i]['prezzo'],DataType::TYPE_NUMERIC);
            $sheet->getCell('U'.$R)->setValueExplicit($margine,DataType::TYPE_FORMULA);
            $sheet->getCell('V'.$R)->setValueExplicit($marginePercentuale,DataType::TYPE_FORMULA);
            $sheet->getCell('W'.$R)->setValueExplicit($margineTotale,DataType::TYPE_FORMULA);
            $sheet->getCell('X'.$R)->setValueExplicit($costoTotale,DataType::TYPE_FORMULA);
            $sheet->getCell('Y'.$R)->setValueExplicit($quantitaTotale,DataType::TYPE_FORMULA);
            $sheet->getCell($colScontoMerceTotale.$R)->setValueExplicit($scontoMerceTotale,DataType::TYPE_FORMULA);

            foreach ($righe[$i]['quantita'] as $quantita) {
				if ($quantita['quantita']) {
                    $col = $colQIndex[$quantita['sede']];
                    $sheet->getCell($col.$R)->setValueExplicit($quantita['quantita'],DataType::TYPE_NUMERIC);
                }

                if ($quantita['scontoMerce']) {
                    $col = $colSCIndex[$quantita['sede']];
                    $sheet->getCell($col.$R)->setValueExplicit($quantita['scontoMerce'],DataType::TYPE_NUMERIC);
                }
            }
        }

        // riquadro di testata (formule)
        // --------------------------------------------------------------------------------
        $totaleMargine = '=SUM('.sprintf("%s%s%s%s%s",'W',$primaRigaDati,':','W',$primaRigaDati+count($righe)-1).')';
        $sheet->getCell('E5')->setValueExplicit($totaleMargine,DataType::TYPE_FORMULA);
        $totaleOrdine = '=SUM('.sprintf("%s%s%s%s%s",'X',$primaRigaDati,':','X',$primaRigaDati+count($righe)-1).')';
        $sheet->getCell('E7')->setValueExplicit($totaleOrdine,DataType::TYPE_FORMULA);
        $rangeI = sprintf("%s%s%s%s%s",'I',$primaRigaDati,':','I',$primaRigaDati+count($righe)-1);
        $rangeX = sprintf("%s%s%s%s%s",'X',$primaRigaDati,':','X',$primaRigaDati+count($righe)-1);
        $marginePercentuale='=IF(E7<>0,ROUND(E5/(SUMPRODUCT('.$rangeI.','.$rangeX.')/100+E7)*100,2),0)';
        $sheet->getCell('E6')->setValueExplicit($marginePercentuale,DataType::TYPE_FORMULA);

        // formattazione
        // --------------------------------------------------------------------------------
        $sheet->getDefaultRowDimension()->setRowHeight(20);
        $sheet->setShowGridlines(true);

        // riquadro di testata
        $sheet->getStyle('B1:C7')->getAlignment()->setHorizontal('left');
        $sheet->getStyle('E1:E7')->getAlignment()->setHorizontal('left');
        $sheet->getStyle('A1:A7')->getFont()->setBold(true);
        $sheet->getStyle('D1:D7')->getFont()->setBold(true);
        //foreach (range('A','X') as $col) {$sheet->getColumnDimension($col)->setAutoSize(true);}

        // colonne descrizione articolo + prezzi
         $sheet->getStyle(sprintf("%s%s%s%s%s",'B',$primaRigaDati,':','C',$primaRigaDati+count($righe)-1))->
            getAlignment()->setHorizontal('center');
        $sheet->getStyle(sprintf("%s%s%s%s%s",'G',$primaRigaDati,':','J',$primaRigaDati+count($righe)-1))->
            getAlignment()->setHorizontal('center');
        $sheet->getStyle(sprintf("%s%s%s%s%s",'L',$primaRigaDati,':','X',$primaRigaDati+count($righe)-1))->
            getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00;');

        // quantita + sconto merce
        $sheet->getStyle(sprintf("%s%s%s%s",$colQuantitaTotale,'1:',$colSCIndex['LAST'],$primaRigaDati-1))->getFont()->setBold(true);
        $sheet->getStyle($colQuantitaTotale.'1')->getAlignment()->setHorizontal('center')->setVertical('center');
        $sheet->getStyle($colScontoMerceTotale.'1')->getAlignment()->setHorizontal('center')->setVertical('center');
        $sheet->getStyle(sprintf("%s%s%s%s%s",$colQuantitaTotale,$primaRigaDati,':',$colSCIndex['LAST'],$primaRigaDati+count($righe)-1))->
            getAlignment()->setHorizontal('center');

        // larghezza colonne (non uso volutamente autowidth)
        $sheet->getColumnDimension('A')->setWidth(25);
        $sheet->getColumnDimension('B')->setWidth(15);
        $sheet->getColumnDimension('C')->setWidth(9);
        $sheet->getColumnDimension('D')->setWidth(28);
        $sheet->getColumnDimension('E')->setWidth(14);
        $sheet->getColumnDimension('F')->setWidth(14);
        $sheet->getColumnDimension('G')->setWidth(10);
        $sheet->getColumnDimension('H')->setWidth(10);
        $sheet->getColumnDimension('I')->setWidth(4);
        $sheet->getColumnDimension('J')->setWidth(4);
        $sheet->getColumnDimension('K')->setWidth(9);
        $sheet->getColumnDimension('L')->setWidth(9);
        $sheet->getColumnDimension('M')->setWidth(6);
        $sheet->getColumnDimension('N')->setWidth(6);
        $sheet->getColumnDimension('O')->setWidth(6);
        $sheet->getColumnDimension('P')->setWidth(6);
        $sheet->getColumnDimension('Q')->setWidth(6);
        $sheet->getColumnDimension('R')->setWidth(6);
        $sheet->getColumnDimension('S')->setWidth(9);
        $sheet->getColumnDimension('T')->setWidth(9);
        $sheet->getColumnDimension('U')->setWidth(9);
        $sheet->getColumnDimension('V')->setWidth(9);
        $sheet->getColumnDimension('W')->setWidth(9);
        $sheet->getColumnDimension('X')->setWidth(9);
        $sheet->getColumnDimension('Y')->setWidth(9);

        $col = $QFirst;
        for ($i = 0; $i<count($filiali); $i++) { //<- quantita
            $sheet->getColumnDimension($col)->setWidth(4);
            $col++;
        }

        $sheet->getColumnDimension($col)->setWidth(9);

        $col = $SCFirst;
        for ($i = 0; $i<count($filiali); $i++) { //<- sconto merce
            $sheet->getColumnDimension($col)->setWidth(4);
            $col++;
        }
        // testata colonne
        $sheet->getStyle('A9:X10')->getAlignment()->setHorizontal('center')->setVertical('center');
        $sheet->getStyle('A9:X10')->getFont()->setBold(true);
        $sheet->getStyle('A9:X10')->getAlignment()->setWrapText(true);
        /*$sheet->getStyle('A9:X10')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle('A9:X10')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle('A9:X10')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle('A9:X10')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THIN);*/

        //$sheet->getStyle('A9:X10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0FFF0');

        $workBook->setActiveSheetIndex(0);
	}

    $writer = new Xlsx($workBook);
    $writer->save($file);

    if (file_exists($file)) {
		header('Content-Description: File Transfer');
		header('Content-Type: application/octet-stream');
		header('Content-Disposition: attachment; filename="'.basename($file).'"');
		header('Expires: 0');
		header('Cache-Control: must-revalidate');
		header('Pragma: public');
		header('Content-Length: ' . filesize($file));
		readfile($file);
		exit;
	}

