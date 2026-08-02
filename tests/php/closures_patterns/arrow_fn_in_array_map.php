<?php
// vybe-test: php/closures_patterns/arrow_fn_in_array_map
// origin: languages/php/tests/php/test_closures_patterns.rs

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

$base = 100;
$result = array_map(fn($n) => $base + $n, [1, 2, 3, 4, 5]);
echo implode(',', $result);

__vybe_check(ob_get_clean(), "101,102,103,104,105");
