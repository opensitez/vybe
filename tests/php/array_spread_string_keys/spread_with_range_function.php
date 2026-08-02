<?php
// vybe-test: php/array_spread_string_keys/spread_with_range_function
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

$prefix = ['start'];
$numbers = [...$prefix, ...range(1, 5)];
echo implode(',', $numbers);

__vybe_check(ob_get_clean(), "start,1,2,3,4,5");
