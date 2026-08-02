<?php
// vybe-test: php/numeric_operations/is_finite_check
// origin: languages/php/tests/php/test_numeric_operations.rs

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

echo is_finite(42.0) ? 'finite' : 'infinite';
echo "\n";
echo is_finite(INF) ? 'finite' : 'infinite';
echo "\n";
echo is_finite(NAN) ? 'finite' : 'infinite';
echo "\n";
echo is_finite(PHP_FLOAT_MAX) ? 'finite' : 'infinite';
echo "\n";

__vybe_check(ob_get_clean(), "finite\ninfinite\ninfinite\nfinite");
