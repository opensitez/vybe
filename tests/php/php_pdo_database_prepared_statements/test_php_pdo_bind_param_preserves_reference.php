<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_bind_param_preserves_reference
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
$pdo->exec("CREATE TABLE t (n INTEGER)");
$stmt = $pdo->prepare("INSERT INTO t VALUES (?)");
$n = 1;
$stmt->bindParam(1, $n, PDO::PARAM_INT);
$stmt->execute();
$n = 2;
$stmt->execute();
$sum = $pdo->query("SELECT SUM(n) FROM t")->fetchColumn();
echo $sum;

__vybe_check(ob_get_clean(), "3");
