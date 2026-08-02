<?php
// vybe-test: php/database/database_class_wrapper
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

class Database {
    public $pdo;
    public function __construct($dsn) {
        $this->pdo = new PDO($dsn);
    }
    public function query($sql) {
        return $this->pdo->query($sql);
    }
    public function exec($sql) {
        return $this->pdo->exec($sql);
    }
}
$db = new Database('sqlite:app.db');
$db->exec('CREATE TABLE IF NOT EXISTS items (id INTEGER PRIMARY KEY, name TEXT)');
$items = $db->query('SELECT * FROM items');
