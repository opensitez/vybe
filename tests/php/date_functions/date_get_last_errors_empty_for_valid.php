<?php
// vybe-test: php/date_functions/date_get_last_errors_empty_for_valid
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

date_parse('2024-12-01');
echo is_array(date_get_last_errors()) ? 'ok' : 'bad';
echo '|';
echo date_get_last_errors()['warning_count'];
echo '|';
echo date_get_last_errors()['error_count'];

__vybe_check(ob_get_clean(), "ok|0|0");
