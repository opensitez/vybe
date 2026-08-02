<?php
// vybe-test: php/type_juggling_strict/is_int_vs_is_float
// origin: languages/php/tests/php/test_type_juggling_strict.rs

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

echo is_int(42) ? 'int' : 'not'; echo is_float(3.14) ? 'float' : 'not';

__vybe_check(ob_get_clean(), "intfloat");
