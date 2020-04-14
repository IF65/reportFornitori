<?php
	@ini_set('memory_limit',-1);

	require(realpath(__DIR__ . '/..').'/vendor/autoload.php');
    //require(realpath(__DIR__ . '/..').'/vendor/PHPMailerAutoload.php');
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
    use PHPMailer\PHPMailer\PHPMailer;
	use Database\Database;

	$timeZone = new \DateTimeZone('Europe/Rome');

    if (file_exists('/huawei') || mkdir (  '/huawei', 0777 , true)) {

        $db = new Database( $sqlDetails );

        $stock = [];

        $dataFine = (new DateTime( 'now', $timeZone ))->sub( new DateInterval( 'P1D' ) );
        $dataInizio = (new DateTime( 'now', $timeZone ))->sub( new DateInterval( 'P60D' ) );
        $periodo = new DatePeriod( $dataInizio, new DateInterval( 'P1D' ), (clone $dataFine)->add( new DateInterval( 'P1D' ) ) );
        foreach ($periodo as $data) {
            $stock[$data->format( 'Y-m-d' )] = $db->huawei->giacenzaAllaData( ['data' => $data->format( 'Y-m-d' )] );
        }

        $response = $db->huawei->caricaDati( ['dataInizio' => $dataInizio->format( 'Y-m-d' ), 'dataFine' => $dataFine->format( 'Y-m-d' )] );

        // creo il file excel
        $style = new Style();

        // creazione del workbook
        $workBook = new Spreadsheet();
        $workBook->getDefaultStyle()->getFont()->setName( 'Arial' );
        $workBook->getDefaultStyle()->getFont()->setSize( 12 );
        $workBook->getProperties()
            ->setCreator( "Supermedia S.p.A. (Gruppo Italmark)" )
            ->setLastModifiedBy( "Supermedia S.p.A." )
            ->setTitle( "Report Vendite Huawei" )
            ->setSubject( "Report Vendite Huawei" )
            ->setDescription( "Report Vendite Huawei" )
            ->setKeywords( "" )
            ->setCategory( "SM Docs" );

        $sheet = $workBook->setActiveSheetIndex( 0 ); // la numerazione dei worksheet parte da 0
        $sheet->setTitle( 'Report vendite Huawei' );

        // testata
        $sheet->getCell( 'A1' )->setValueExplicit( 'anno', DataType::TYPE_STRING );
        $sheet->getCell( 'B1' )->setValueExplicit( 'settimana', DataType::TYPE_STRING );
        $sheet->getCell( 'C1' )->setValueExplicit( 'data', DataType::TYPE_STRING );
        $sheet->getCell( 'D1' )->setValueExplicit( 'cliente', DataType::TYPE_STRING );
        $sheet->getCell( 'E1' )->setValueExplicit( 'negozio', DataType::TYPE_STRING );
        $sheet->getCell( 'F1' )->setValueExplicit( 'negozioDescrizione', DataType::TYPE_STRING );
        $sheet->getCell( 'G1' )->setValueExplicit( 'barcode', DataType::TYPE_STRING );
        $sheet->getCell( 'H1' )->setValueExplicit( 'codice', DataType::TYPE_STRING );
        $sheet->getCell( 'I1' )->setValueExplicit( 'descrizione', DataType::TYPE_STRING );
        $sheet->getCell( 'J1' )->setValueExplicit( 'quantita', DataType::TYPE_STRING );
        $sheet->getCell( 'K1' )->setValueExplicit( 'stock', DataType::TYPE_STRING );
        $sheet->getCell( 'L1' )->setValueExplicit( 'importoTotale', DataType::TYPE_STRING );

        foreach ($response['data'] as $rowIndex => $row) {

            $giacenza = 0;
            if (array_key_exists( $row['data'], $stock )) {
                $stockData = $stock[$row['data']];
                if (array_key_exists( $row['codice'], $stockData )) {
                    $stockCodice = $stockData[$row['codice']];
                    if (array_key_exists( $row['negozio'], $stockCodice )) {
                        $giacenza = $stockCodice[$row['negozio']];
                    }
                }
            }

            $r = sprintf( '%d', $rowIndex + 2 );
            $sheet->getCell( 'A' . $r )->setValueExplicit( $row['anno'], DataType::TYPE_NUMERIC );
            $sheet->getCell( 'B' . $r )->setValueExplicit( $row['settimana'], DataType::TYPE_NUMERIC );
            $sheet->getCell( 'C' . $r )->setValueExplicit( $row['data'], DataType::TYPE_STRING );
            $sheet->getCell( 'D' . $r )->setValueExplicit( $row['cliente'], DataType::TYPE_STRING );
            $sheet->getCell( 'E' . $r )->setValueExplicit( $row['negozio'], DataType::TYPE_STRING );
            $sheet->getCell( 'F' . $r )->setValueExplicit( $row['negozio_descrizione'], DataType::TYPE_STRING );
            $sheet->getCell( 'G' . $r )->setValueExplicit( $row['barcode'], DataType::TYPE_STRING );
            $sheet->getCell( 'H' . $r )->setValueExplicit( $row['codice'], DataType::TYPE_STRING );
            $sheet->getCell( 'I' . $r )->setValueExplicit( $row['descrizione'], DataType::TYPE_STRING );
            $sheet->getCell( 'J' . $r )->setValueExplicit( $row['quantita'], DataType::TYPE_NUMERIC );
            $sheet->getCell( 'K' . $r )->setValueExplicit( $giacenza, DataType::TYPE_NUMERIC );
            $sheet->getCell( 'L' . $r )->setValueExplicit( $row['importoTotale'], DataType::TYPE_NUMERIC );
        }

        $fileName  = '/huawei/huawei_' . $dataFine->format( 'Ymd' ) . '.xlsx';
        $writer = new Xlsx( $workBook );
        $writer->save( $fileName );

        $connesso = false;

        $connId = ftp_connect( '80.158.18.163' );
        if ($connId) {
            if (@ftp_login( $connId, 'SUPERMEDIA', 'PpF4A!j!HYe!rThM' )) {
                if (ftp_put( $connId, "huawei.xlsx", $fileName, FTP_BINARY )) {
                    rename($fileName,'/huawei/ok_huawei_' . $dataFine->format( 'Ymd' ) . '.xlsx');
                }
                ftp_close( $connId );
            }
        }

        /*$connId = ftp_connect( '10.11.14.78' );
        if ($connId) {
            if (@ftp_login( $connId, 'root', 'BGT567ujm' )) {
                if (ftp_put( $connId, "huawei.xlsx", $fileName, FTP_BINARY )) {
                    rename($fileName,'/huawei/ok_huawei_' . $dataFine->format( 'Ymd' ) . '.xlsx');
                }
                ftp_close( $connId );
            }
        }*/

        $mail = new PHPMailer;

        $mail->SMTPOptions = array(
            'ssl' => array(
                'verify_peer' => false,
                'verify_peer_name' => false,
                'allow_self_signed' => true
            )
        );

        $mail->isSMTP();                                      // Set mailer to use SMTP
        $mail->Host = '10.11.14.233:25';  // Specify main and backup SMTP servers
        //$mail->SMTPAuth = true;                               // Enable SMTP authentication
        //$mail->Username = 'marco.gnecchi@if65.it';                 // SMTP username
        //$mail->Password = 'MGnecchi1';                           // SMTP password
        //$mail->SMTPSecure = 'tls';                            // Enable encryption, 'ssl' also accepted

        $mail->From = 'edp@supermedia.it';
        $mail->FromName = 'EDP Supermedia';
        $mail->addAddress('marco.gnecchi@supermedia.it', 'Marco Gnecchi');     // Add a recipient
        //$mail->addAddress('nicola.pirovano@supermedia.it', 'Nicola Pirovano');     // Add a recipient
        //$mail->addAddress('sergio.guidi@supermedia.it', 'Sergio Guidi');     // Add a recipient
        //$mail->addAddress('ellen@example.com');               // Name is optional
        $mail->addReplyTo('marco.gnecchi@supermedia.it', 'Marco Gnecchi');
        $mail->addCC('marco.gnecchi@supermedia.it');
        //$mail->addBCC('bcc@example.com');

        $mail->WordWrap = 50;                                 // Set word wrap to 50 characters
        $mail->addAttachment('/huawei/ok_huawei_' . $dataFine->format( 'Ymd' ) . '.xlsx');         // Add attachments
        //$mail->addAttachment('/tmp/image.jpg', 'new.jpg');    // Optional name
        $mail->isHTML(true);                                  // Set email format to HTML

        $mail->Subject = 'Invio dati Huawei ' . $dataFine->format( 'd/m/Y' );
        $mail->Body    = '<b>Invio venduto Huawei del '.$dataFine->format( 'd/m/Y' )."</b>\n";
        //$mail->AltBody = 'This is the body in plain text for non-HTML mail clients';

        $mail->send();

    }