<?php
// vybe-test: php/array_key_destructuring/foreach_list_with_reference_updates_source
// origin: languages/php/tests/php/test_array_key_destructuring.rs

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

$pairs = [[1], [2], [3]];
$sum = 0;
foreach ($pairs as &$pair) {
    $pair[0] *= 10;
    $sum += $pair[0];
}
echo $sum;

__vybe_check(ob_get_clean(), "60");
