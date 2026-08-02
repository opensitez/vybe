<?php
// vybe-test: php/dateinterval_format_specifiers/dateinterval_format_specifiers
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

$i = new DateInterval('P2Y4DT6H8M');
echo $i->format('%y years, %d days, %h hours, %i minutes');

__vybe_check(ob_get_clean(), "2 years, 4 days, 6 hours, 8 minutes");
