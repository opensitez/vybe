<?php
// vybe-test: php/array_spread_string_keys/spread_of_filtered_array_reindexes
// origin: languages/php/tests/php/test_array_spread_string_keys.rs

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

$filtered = array_values(array_filter([1, 2, 3, 4, 5], fn($x) => $x % 2 === 0));
$result = [0, ...$filtered, 6];
echo implode(',', $result);

__vybe_check(ob_get_clean(), "0,2,4,6");
