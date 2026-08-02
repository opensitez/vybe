<?php
// vybe-test: php/numeric_operations/php_float_epsilon_and_max
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

echo (PHP_FLOAT_EPSILON > 0) ? 'positive' : 'zero';
echo "\n";
echo (PHP_FLOAT_MAX > 1e100) ? 'large' : 'small';
echo "\n";
echo PHP_FLOAT_DIG . "\n";

__vybe_check(ob_get_clean(), "positive\nlarge\n15");
