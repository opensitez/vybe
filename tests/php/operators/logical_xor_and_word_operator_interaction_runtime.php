<?php
// vybe-test: php/operators/logical_xor_and_word_operator_interaction_runtime
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

echo (true xor false) ? 'x1' : 'x0';
echo '|';
echo (true and false) ? 'and1' : 'and0';
echo '|';
echo (true or false) ? 'or1' : 'or0';

__vybe_check(ob_get_clean(), "x1|and0|or1");
