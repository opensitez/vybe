<?php
// vybe-test: php/dateinterval_format_specifiers/dateinterval_format_full_pattern
// origin: languages/php/tests/php/test_dateinterval_format_specifiers.rs

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

$i = new DateInterval('P3Y2M1DT4H5M6S');
echo $i->format('%Y years %M months %D days %H hours %I minutes %S seconds');

__vybe_check(ob_get_clean(), "03 years 02 months 01 days 04 hours 05 minutes 06 seconds");
