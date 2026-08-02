<?php
// vybe-test: php/database/mysqli_adapter_exposes_connect_error_helpers
// origin: languages/php/tests/php/test_database.rs

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

mysqli_report(0);
$dbh = mysqli_init();
mysqli_real_connect($dbh, 'localhost', 'user', 'pass', null, null, null, 0);
echo mysqli_connect_errno();
echo mysqli_connect_error();
echo mysqli_error($dbh);

__vybe_check(ob_get_clean(), "1\nConnection failed\nConnection failed");
