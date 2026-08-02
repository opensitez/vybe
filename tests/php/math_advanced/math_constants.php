<?php
// vybe-test: php/math_advanced/math_constants
// origin: languages/php/tests/php/test_math_advanced.rs

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

echo round(M_PI, 5);
echo round(M_E, 5);
echo round(M_SQRT2, 5);
echo round(M_LN2, 5);
echo round(M_LOG2E, 5);

__vybe_check(ob_get_clean(), "3.141592.718281.414210.693151.4427");
