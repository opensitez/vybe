<?php
// vybe-test: php/datetime_parsing/datetime_get_last_errors_clean_after_valid_create
// origin: languages/php/tests/php/test_datetime_parsing.rs

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

DateTime::createFromFormat('Y-m-d', '2024-06-01');
$e = DateTime::getLastErrors();
echo (($e['error_count'] ?? 0) + ($e['warning_count'] ?? 0)) === 0 ? 'clean' : 'dirty';

__vybe_check(ob_get_clean(), "clean");
