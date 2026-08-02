<?php
// vybe-test: php/mysqli_prepared/mysqli_prepare_returns_stmt_object
// origin: languages/php/tests/php/test_mysqli_prepared.rs

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

$mysqli = new mysqli('127.0.0.1', 'u', 'p', 'db');
$stmt = $mysqli->prepare('SELECT 1');
echo $stmt === false ? 'fail' : get_class($stmt);

__vybe_check(ob_get_clean(), "mysqli_stmt");
