<?php
// vybe-test: php/math_functions/math_round_half_even_runtime
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

echo round(2.5);
echo '|';
echo round(3.5, 0, PHP_ROUND_HALF_UP);

__vybe_check(ob_get_clean(), "2|4");
