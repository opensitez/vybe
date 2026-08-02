<?php
// vybe-test: php/php_pdo_sqlite_transactions_savepoints/test_php_pdo_fetch_column_indexed_column
// origin: languages/php/tests/php/test_php_pdo_sqlite_transactions_savepoints.rs
// vybe-test-mode: compile

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE t (a INT, b TEXT)");
$pdo->exec("INSERT INTO t VALUES (10, 'ten')");
$stmt = $pdo->query("SELECT a, b FROM t");
echo "B_VAL=" . $stmt->fetchColumn(1);
