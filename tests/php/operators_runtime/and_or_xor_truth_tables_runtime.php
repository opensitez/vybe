<?php
// vybe-test: php/operators_runtime/and_or_xor_truth_tables_runtime
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

echo (0 and 0) ? 'a' : 'f';
echo '|';
echo (0 and 1) ? 'a' : 'f';
echo '|';
echo (1 or 0) ? 'a' : 'f';
echo '|';
echo (1 xor 1) ? 'a' : 'f';
echo '|';
echo (1 xor 0) ? 'a' : 'f';

__vybe_check(ob_get_clean(), "f|f|a|f|a");
