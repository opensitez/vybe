<?php
// vybe-test: php/misc_builtins/mktime_returns_timestamp
// origin: languages/php/tests/php/test_misc_builtins.rs

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

echo mktime(0,0,0,1,1,1970) >= 0 ? 'pos_or_zero' : 'neg';

__vybe_check(ob_get_clean(), "pos_or_zero");
