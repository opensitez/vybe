<?php
// vybe-test: php/pdo_fetch_func_callback/pdo_fetch_func_callback
// origin: languages/php/tests/php/test_pdo_fetch_func_callback.rs

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

$pdo = new PDO('sqlite::memory:');
$pdo->exec("CREATE TABLE numbers (a INTEGER, b INTEGER)");
$pdo->exec("INSERT INTO numbers VALUES (5, 10)");
$pdo->exec("INSERT INTO numbers VALUES (3, 7)");

$stmt = $pdo->query("SELECT a, b FROM numbers");
$results = $stmt->fetchAll(PDO::FETCH_FUNC, function($a, $b) {
    return $a + $b;
});

echo implode(',', $results);

__vybe_check(ob_get_clean(), "15,10");
