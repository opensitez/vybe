<?php
// vybe-test: php/php_bitwise_shifts_flags_masking/test_php_bitwise_and_or_xor_not_primitives
// origin: languages/php/tests/php/test_php_bitwise_shifts_flags_masking.rs

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

$a = 0b1100; // 12
$b = 0b1010; // 10

$and = $a & $b; // 1000 (8)
$or = $a | $b;  // 1110 (14)
$xor = $a ^ $b; // 0110 (6)

echo "$and | $or | $xor";

__vybe_check(ob_get_clean(), "8 | 14 | 6");
