<?php
// vybe-test: php/php_pdo_sqlite_transactions_savepoints/test_php_pdo_prepare_positional_parameters
// origin: languages/php/tests/php/test_php_pdo_sqlite_transactions_savepoints.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE scores (player TEXT, score INT)");
$stmt = $pdo->prepare("INSERT INTO scores VALUES (?, ?)");
$stmt->execute(["Alice", 100]);
$stmt->execute(["Bob", 200]);

$query = $pdo->query("SELECT SUM(score) FROM scores");
echo "Total Score: " . $query->fetchColumn();
