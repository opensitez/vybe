<?php
// vybe-test: php/operators_runtime/bitwise_shift_then_bitwise_runtime
// origin: languages/php/tests/php/test_operators_runtime.rs

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

echo (1 << 3) & 14;
echo '|';
echo 1 << (3 & 2);
echo '|';
echo (8 >> 1) ^ 3;
echo '|';
echo (9 >> 1) | 2;

__vybe_check(ob_get_clean(), "8|4|7|6");
