<?php
// vybe-test: php/date_functions/date_parse_invalid_input_reports_error_count
// origin: languages/php/tests/php/test_date_functions.rs

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

$parsed = date_parse('not-a-date');
echo is_array($parsed) && ($parsed['error_count'] ?? 0) > 0 ? 'err' : 'ok';

__vybe_check(ob_get_clean(), "err");
