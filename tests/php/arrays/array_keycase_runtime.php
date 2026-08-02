<?php
// vybe-test: php/arrays/array_keycase_runtime
// origin: languages/php/tests/php/test_arrays.rs

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

$a = ['One' => 1, 'Two' => 2];
$lower = array_change_key_case($a, CASE_LOWER);
$upper = array_change_key_case($a, CASE_UPPER);
echo isset($lower['one']) ? '1' : '0';
echo isset($upper['TWO']) ? '|1' : '|0';

__vybe_check(ob_get_clean(), "1|1");
