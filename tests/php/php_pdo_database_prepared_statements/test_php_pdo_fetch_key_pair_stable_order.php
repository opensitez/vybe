<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_fetch_key_pair_stable_order
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
$pdo->exec("CREATE TABLE map (k TEXT, v TEXT)");
$pdo->exec("INSERT INTO map VALUES ('a', '1'), ('b', '2'), ('c', '3')");
$rows = $pdo->query("SELECT k, v FROM map")->fetchAll(PDO::FETCH_KEY_PAIR);
ksort($rows);
echo implode('|', array_keys($rows));

__vybe_check(ob_get_clean(), "a|b|c");
