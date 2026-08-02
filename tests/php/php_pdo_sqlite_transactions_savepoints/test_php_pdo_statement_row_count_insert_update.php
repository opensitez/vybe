<?php
// vybe-test: php/php_pdo_sqlite_transactions_savepoints/test_php_pdo_statement_row_count_insert_update
// origin: languages/php/tests/php/test_php_pdo_sqlite_transactions_savepoints.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE items (id INT, val TEXT)");
$stmt = $pdo->prepare("INSERT INTO items VALUES (?, ?)");
$stmt->execute([1, "a"]);
echo "Affected: " . $stmt->rowCount();
