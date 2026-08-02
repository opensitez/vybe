<?php
// vybe-test: php/bitwise_operators/bitwise_chained_flags_state_mutation_runtime
// origin: languages/php/tests/php/test_bitwise_operators.rs

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

$state = 0;
$state |= 0b0001;
$state |= 0b0100;
echo decbin($state);
$state &= ~0b0001;
echo '|';
echo decbin($state);
$state ^= 0b0110;
echo '|';
echo decbin($state);

__vybe_check(ob_get_clean(), "101|100|10");
