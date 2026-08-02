<?php
// vybe-test: php/array_map_multiple/array_filter_with_strict_callable_on_empty_array
// origin: languages/php/tests/php/test_array_map_multiple.rs

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

$values = [];
$filtered = array_filter($values, fn($x) => $x > 1);
echo is_array($filtered) ? 'array' : 'no';
echo '|';
echo count($filtered);

__vybe_check(ob_get_clean(), "array|0");
