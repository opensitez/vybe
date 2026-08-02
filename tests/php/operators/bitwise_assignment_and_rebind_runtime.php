<?php
// vybe-test: php/operators/bitwise_assignment_and_rebind_runtime
// origin: languages/php/tests/php/test_operators.rs

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

$bits = 0b1111;
$bits &= 0b1010;
echo $bits;
echo '|';
$bits |= 0b0101;
echo $bits;
echo '|';
$bits ^= 0b0011;
echo $bits;
$bits <<= 1;
echo '|';
echo $bits;
$bits >>= 2;
echo '|';
echo $bits;

__vybe_check(ob_get_clean(), "10|15|12|24|6");
