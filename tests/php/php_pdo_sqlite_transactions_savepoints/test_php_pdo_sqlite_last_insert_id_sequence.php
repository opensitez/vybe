<?php
// vybe-test: php/php_pdo_sqlite_transactions_savepoints/test_php_pdo_sqlite_last_insert_id_sequence
// origin: languages/php/tests/php/test_php_pdo_sqlite_transactions_savepoints.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT)");
$pdo->exec("INSERT INTO users (username) VALUES ('Alice')");
$id1 = $pdo->lastInsertId();
$pdo->exec("INSERT INTO users (username) VALUES ('Bob')");
$id2 = $pdo->lastInsertId();

echo "ID1=$id1 ID2=$id2";
