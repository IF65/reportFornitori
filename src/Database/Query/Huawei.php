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
                $dallaData = $query['dallaData'];
                $allaData = $query['allaData'];
                
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
                            0 `stock`,
                            r.`importo_totale` `importoTotale`
                        from db_sm.righe_vendita as r left join archivi.negozi as n on n.codice_interno = r.`negozio`
                        where r.`data` >= '2018-08-01' and (r.`linea` = 'HUAWEI TEL' or r.`linea` = 'HUAWEI TAB') and r.`riparazione` = 0 and
                            r.`codice` not in ('0501754','0501763','0538797','0538804','0538831','0538840','0546332','0556151','0556160')";
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

        public function __destruct() {
			unset($this->pdo);
        }

    }
?>
