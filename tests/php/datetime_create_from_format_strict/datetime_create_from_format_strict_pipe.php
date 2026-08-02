<?php
// vybe-test: php/datetime_create_from_format_strict/datetime_create_from_format_strict_pipe
// origin: languages/php/tests/php/test_datetime_create_from_format_strict.rs

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

// The '|' resets all fields to the Unix Epoch
$dt = DateTime::createFromFormat('Y-m-d|', '2009-02-15');
echo $dt->format('Y-m-d H:i:s');

__vybe_check(ob_get_clean(), "2009-02-15 00:00:00");
