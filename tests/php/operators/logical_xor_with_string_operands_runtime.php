<?php
// vybe-test: php/operators/logical_xor_with_string_operands_runtime
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

echo ((bool)'x' xor (bool) '') ? 'xor1' : 'xor0';
echo '|';
echo ((bool)'x' xor (bool)'y') ? 'xor1' : 'xor0';

__vybe_check(ob_get_clean(), "xor1|xor0");
