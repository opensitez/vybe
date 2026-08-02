<?php
// vybe-test: php/operators/unary_plus_and_unary_minus_runtime_values
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

echo +5 . '|';
echo +(-7) . '|';
echo +('8') . '|';
echo -(-3) . '|';
echo -(1 + 2) . '|';
echo +3.5;

__vybe_check(ob_get_clean(), "5|-7|8|3|-3|3.5");
