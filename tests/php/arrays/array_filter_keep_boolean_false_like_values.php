<?php
// vybe-test: php/arrays/array_filter_keep_boolean_false_like_values
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

$values = [0, '0', '', null, false, 2, 3];
$a = array_filter($values, fn($v) => $v !== null && $v !== false);
$b = array_filter($values, fn($v) => is_int($v));
echo count($a) . '|';
echo count($b);

__vybe_check(ob_get_clean(), "5|3");
