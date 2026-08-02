<?php
// vybe-test: php/bitwise_shifts_runtime/xor_swap_without_temp
// origin: languages/php/tests/php/test_bitwise_shifts_runtime.rs

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

$a = 5;
$b = 9;
$a ^= $b;
$b ^= $a;
$a ^= $b;
echo $a . $b;

__vybe_check(ob_get_clean(), "95");
