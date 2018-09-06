<?php
    namespace Database;

	use \PDO;
    use Database\Query\Huawei;

    class Database {

        protected $pdo = null;
        
        public $huawei;

        public function __construct($sqlDetails) {
            $conStr = sprintf("mysql:host=%s", $sqlDetails['host']);
            try {
                $this->pdo = new PDO($conStr, $sqlDetails['user'], $sqlDetails['password']);
                
                $this->huawei = new Huawei($this->pdo);

            } catch (PDOException $e) {
                die($e->getMessage());
            }
        }
        
        public function estrazioneDatiHuawei(array $query) {
            $report = $this->tableAnagdafi->ricerca($query);
            
            return (array) $report;
        }
        
        public function __destruct() {
            $this->pdo = null;
        }
    }
?>
