<?php
// vybe-test: php/types/empty_values
// origin: languages/php/tests/php/test_types.rs

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

echo empty('') ? 't' : 'f'; echo empty(0) ? 't' : 'f'; echo empty(null) ? 't' : 'f'; echo empty('0') ? 't' : 'f'; echo empty('x') ? 't' : 'f'; echo empty([]) ? 't' : 'f';

__vybe_check(ob_get_clean(), "ttttft");
