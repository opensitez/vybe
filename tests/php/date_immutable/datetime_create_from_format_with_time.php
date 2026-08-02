<?php
// vybe-test: php/date_immutable/datetime_create_from_format_with_time
// origin: languages/php/tests/php/test_date_immutable.rs

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

$d = DateTime::createFromFormat('Y-m-d H:i', '2024-06-15 14:30');
echo $d->format('H:i');

__vybe_check(ob_get_clean(), "14:30");
