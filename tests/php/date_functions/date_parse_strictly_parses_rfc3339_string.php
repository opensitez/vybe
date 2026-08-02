<?php
// vybe-test: php/date_functions/date_parse_strictly_parses_rfc3339_string
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

$dt = date_parse('2024-07-01T15:30:00Z');
echo $dt['year'] . '-' . str_pad((string) $dt['month'], 2, '0', STR_PAD_LEFT) . '-' . str_pad((string) $dt['day'], 2, '0', STR_PAD_LEFT);

__vybe_check(ob_get_clean(), "2024-07-01");
