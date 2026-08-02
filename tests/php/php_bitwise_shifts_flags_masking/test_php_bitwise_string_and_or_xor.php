<?php
// vybe-test: php/php_bitwise_shifts_flags_masking/test_php_bitwise_string_and_or_xor
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

$s1 = "A"; // ASCII 65 (01000001)
$s2 = " "; // ASCII 32 (00100000)
$res = $s1 | $s2; // ASCII 97 ('a')
echo $res;

__vybe_check(ob_get_clean(), "a");
