<?php
// vybe-test: php/database/pdo_transaction
// origin: languages/php/tests/php/test_database.rs
// vybe-test-mode: compile

$pdo = new PDO('sqlite:test.db');
$pdo->beginTransaction();
$pdo->exec("INSERT INTO users (name) VALUES ('Bob')");
$pdo->commit();
