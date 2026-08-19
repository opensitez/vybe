<?php
// vybe-test: php/php_math_bcmath_number_precision/test_php_bcsqrt_and_bcpow
// origin: languages/php/tests/php/test_php_math_bcmath_number_precision.rs

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

$sqrt = bcsqrt("2", 6);
$pow = bcpow("2", "10", 0);
echo "sqrt2=$sqrt pow=$pow";


__vybe_check(ob_get_clean(), "sqrt2=1.414213 pow=1024");
