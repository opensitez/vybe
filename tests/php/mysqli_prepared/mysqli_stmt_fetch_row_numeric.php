<?php
// vybe-test: php/mysqli_prepared/mysqli_stmt_fetch_row_numeric
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

$stmt = (new mysqli())->prepare('SELECT 1 AS n');
$stmt->execute();
$row = $stmt->get_result()->fetch_row();
echo $row[0] ?? 'x';

__vybe_check(ob_get_clean(), "1");
