<?php
// vybe-test: php/math_advanced/modf_integer_and_fraction
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

[$frac, $int] = modf(3.25);
echo $frac . "\n";
echo $int . "\n";
[$frac2, $int2] = modf(-7.75);
echo $frac2 . "\n";
echo $int2 . "\n";

__vybe_check(ob_get_clean(), "0.25\n3\n-0.75\n-7");
