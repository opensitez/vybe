<?php
// vybe-test: php/php_pdo_sqlite_transactions_savepoints/test_php_pdo_in_transaction_status_flag
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
echo $pdo->inTransaction() ? "TX1" : "TX0";
$pdo->beginTransaction();
echo $pdo->inTransaction() ? " TX1" : " TX0";
$pdo->commit();
echo $pdo->inTransaction() ? " TX1" : " TX0";

__vybe_check(ob_get_clean(), "TX0 TX1 TX0");
