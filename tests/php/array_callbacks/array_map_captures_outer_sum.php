<?php
// vybe-test: php/array_callbacks/array_map_captures_outer_sum
// origin: languages/php/tests/php/test_array_callbacks.rs

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

$numbers = [1, 2, 3];
$base = 10;
$mapped = array_map(fn($n) => $n + $base, $numbers);
echo implode(',', $mapped);

__vybe_check(ob_get_clean(), "11,12,13");
