<?php
// vybe-test: php/programs/base_conversion_decimal_to_binary
// origin: languages/php/tests/php/test_programs.rs

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

function toBinary(int $n): string {
    if ($n === 0) return '0';
    $bits = '';
    while ($n > 0) { $bits = ($n % 2) . $bits; $n = intdiv($n, 2); }
    return $bits;
}
echo toBinary(0) . "\n";
echo toBinary(10) . "\n";
echo toBinary(255) . "\n";

__vybe_check(ob_get_clean(), "0\n1010\n11111111");
