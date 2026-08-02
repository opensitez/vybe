<?php
// vybe-test: php/math_functions/fdiv_zero_division_behavior_runtime
// origin: languages/php/tests/php/test_math_functions.rs

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

echo fdiv(5.0, 0.0);
echo '|';
echo fdiv(-5.0, 0.0);
echo '|';
echo is_infinite(fdiv(1, 0)) ? '1' : '0';

__vybe_check(ob_get_clean(), "INF|-INF|1");
