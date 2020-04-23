<?php
    namespace Database\Query;

	class Huawei {
        private $pdo = null;

        public function __construct($pdo) {
        	try {
                $this->pdo = $pdo;
                
            } catch (PDOException $e) {
                die($e->getMessage());
            }
        }

        public function caricaDati(array $query) {
            try {
                $dataInizio = $query['dataInizio'];
                $dataFine = $query['dataFine'];
                
                $sql = "select 
                            year(r.`data`) `anno`,
                            week(r.`data`, 1) `settimana`,
                            r.`data`,
                            'SUPERMEDIA' `cliente`,
                            r.`negozio`,
                            ifnull(n.`negozio_descrizione`,'DA DEFINIRE') `negozio_descrizione`,
                            case when r.`ean`<>'' then r.`ean` else (select ifnull(e.ean,'NULLO') from db_sm.`ean` as e where e.codice = r.`codice` limit 1) end `barcode`,
                            r.`codice`,
                            r.`descrizione`,
                            r.`quantita`,
                            case when s.`giacenza` < 0 then 0 else s.`giacenza` end  `stock`,
                            r.`importo_totale` `importoTotale`
                        from db_sm.righe_vendita as r left join archivi.negozi as n on n.codice_interno = r.`negozio` left join db_sm.situazioni as s on r.`negozio`=s.`negozio` and r.`codice`=s.`codice_articolo`
                        where r.`data` >= '$dataInizio' and r.`data` <= ' $dataFine' and (r.`linea` like 'HUAWEI%' or r.`linea`='HONOR TEL') and r.`riparazione` = 0 and
                            r.`codice` not in ('0501754','0501763','0538797','0538804','0538831','0538840','0546332','0556151','0556160','0678822','0686341','0686350','0765906','0765915','0765924')
                        order by r.`data`,r.`codice`,r.`negozio`";
                
                $data = [];
                $stmt = $this->pdo->prepare( $sql );
                $stmt->execute();
                while ($row = $stmt->fetch(\PDO::FETCH_ASSOC, \PDO::FETCH_ORI_NEXT)) {
                    $data[] = $row;
                }
                $stmt = null;
                
                return array("recordsTotal" => count($data) ,"data" => $data);
            } catch (PDOException $e) {
                die($e->getMessage());
            }
        }
        
        public function giacenzaAllaData(array $query) {
            $data = $query['data'];
            $sql = "select
                        g.codice,
                        g.negozio,
                        g.giacenza
                    from db_sm.giacenze as g join 
                        (
                            select
                                codice,
                                negozio,
                                max(data) data
                            from db_sm.giacenze
                            where data = '$data' and  codice in
                            (
                                select
                                    codice
                                    from db_sm.magazzino
                                where linea in ('HUAWEI TEL','HUAWEI TAB','HONOR TEL') and negozio <> 'SMMD'
                            ) group by codice, negozio order by data desc
                        ) as s on g.data = s.data and g.negozio=s.negozio and g.codice = s.codice";
                        
                        
            $stock = [];
            $stmt = $this->pdo->prepare( $sql );
            $stmt->execute();
            while ($row = $stmt->fetch(\PDO::FETCH_ASSOC, \PDO::FETCH_ORI_NEXT)) {
                if (array_key_exists($row['codice'],$stock)) {
                   $stock[$row['codice']][$row['negozio']] = $row['giacenza'] * 1;
                } else {
                    $stock[$row['codice']] = [$row['negozio'] => $row['giacenza'] * 1];
                }
                
            }
            $stmt = null;
                
            return $stock;
        }

        public function __destruct() {
			unset($this->pdo);
        }

    }
    
    /*
     *
     *  ftp://80.158.18.163
        Port: 21
        Root:\

        Supermedia FTP
        ID: SUPERMEDIA
        PW: PpF4A!j!HYe!rThM
     *
     *
     *
     */
?>
