<?php
// vybe-test: php/datetime/datetime_localtime_fields_runtime
// origin: languages/php/tests/php/test_datetime.rs

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

$ts = mktime(13, 5, 7, 2, 3, 2024);
$lt = localtime($ts, true);
echo $lt['tm_mday'];
echo $lt['tm_mon'];
echo $lt['tm_year'] + 1900;

__vybe_check(ob_get_clean(), "32024");
