<?php
// vybe-test: php/php_pdo_sqlite_transactions_savepoints/test_php_pdo_nested_transactions_savepoints_sqlite
// origin: languages/php/tests/php/test_php_pdo_sqlite_transactions_savepoints.rs

function __vybe_check($got, $want) {
    // Match the Rust harness's normalisation: strip \r, then drop trailing
    // newlines (it split on "\n" and popped empty trailing elements).
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    // Replay the program's own output so running the file by hand still
    // behaves like the program it was extracted from.
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

$pdo = new PDO("sqlite::memory:");
$pdo->exec("CREATE TABLE logs (msg TEXT)");

$pdo->beginTransaction();
$pdo->exec("INSERT INTO logs VALUES ('Outer 1')");

$pdo->exec("SAVEPOINT sp1");
$pdo->exec("INSERT INTO logs VALUES ('Inner 1')");
$pdo->exec("ROLLBACK TO SAVEPOINT sp1");

$pdo->commit();

$stmt = $pdo->query("SELECT group_concat(msg, ',') FROM logs");
echo $stmt->fetchColumn();

__vybe_check(ob_get_clean(), "Outer 1");
