<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_column_count_and_error_code_runtime
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
$stmt = $pdo->prepare("SELECT * FROM t");
$stmt->execute();
echo $stmt->columnCount();

@$stmt2 = $pdo->query("SELECT * FROM nope");
echo $pdo->errorCode();

__vybe_check(ob_get_clean(), "0\n42S02");
