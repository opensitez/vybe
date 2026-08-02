<?php
// vybe-test: php/numeric_operations/pi_constant
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

echo round(pi(), 5) . "\n";
echo round(M_PI, 5) . "\n";
echo (pi() === M_PI) ? 'equal' : 'not equal';
echo "\n";

__vybe_check(ob_get_clean(), "3.14159\n3.14159\nequal");
