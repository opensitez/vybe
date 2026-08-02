<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_transaction_commit_and_rollback
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs

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
$pdo->exec("CREATE TABLE accounts (id INT, balance REAL)");
$pdo->exec("INSERT INTO accounts VALUES (1, 100.0)");

try {
    $pdo->beginTransaction();
    $pdo->exec("UPDATE accounts SET balance = balance - 50.0 WHERE id = 1");
    $pdo->rollBack();
} catch (Exception $e) {}

$stmt = $pdo->query("SELECT balance FROM accounts WHERE id = 1");
echo "Balance: " . $stmt->fetchColumn();

__vybe_check(ob_get_clean(), "Balance: 100");
