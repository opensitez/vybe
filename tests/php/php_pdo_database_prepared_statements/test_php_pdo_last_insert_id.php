<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_last_insert_id
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE logs (id INTEGER PRIMARY KEY AUTOINCREMENT, msg TEXT)");
$pdo->exec("INSERT INTO logs (msg) VALUES ('log1')");
echo "ID: " . $pdo->lastInsertId();
