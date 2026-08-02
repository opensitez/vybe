<?php
// vybe-test: php/php84_datetime_create_from_timestamp/test_php84_datetime_mutable_create_from_timestamp
// origin: languages/php/tests/php/test_php84_datetime_create_from_timestamp.rs

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

$ts = 1600000000;
if (method_exists('DateTime', 'createFromTimestamp')) {
    $dt = DateTime::createFromTimestamp($ts);
    echo "Mutable: " . $dt->format("Y-m-d");
} else {
    $dt = (new DateTime())->setTimestamp($ts);
    echo "Mutable: " . $dt->format("Y-m-d");
}

__vybe_check(ob_get_clean(), "Mutable: 2020-09-13");
