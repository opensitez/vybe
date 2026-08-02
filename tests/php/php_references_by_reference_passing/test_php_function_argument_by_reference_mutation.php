<?php
// vybe-test: php/php_references_by_reference_passing/test_php_function_argument_by_reference_mutation
// origin: languages/php/tests/php/test_php_references_by_reference_passing.rs

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

function increment(int &$num, int $step = 1): void {
    $num += $step;
}

$val = 5;
increment($val, 10);
echo $val;

__vybe_check(ob_get_clean(), "15");
