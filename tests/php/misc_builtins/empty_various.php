<?php
// vybe-test: php/misc_builtins/empty_various
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

echo empty('') ? '1' : '0';
echo empty(0) ? '1' : '0';
echo empty([]) ? '1' : '0';
echo empty('hello') ? '1' : '0';
echo empty(1) ? '1' : '0';

__vybe_check(ob_get_clean(), "11100");
