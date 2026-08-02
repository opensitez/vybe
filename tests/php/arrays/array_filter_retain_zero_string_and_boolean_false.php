<?php
// vybe-test: php/arrays/array_filter_retain_zero_string_and_boolean_false
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

$values = [0, 0.0, '0', false, true, 'ok', ''];
$strict = array_filter($values, fn($v) => $v !== null);
$loose = array_filter($values);
echo count($strict) . '|';
echo count($loose);

__vybe_check(ob_get_clean(), "7|2");
