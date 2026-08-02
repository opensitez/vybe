<?php
// vybe-test: php/array_map_multiple/array_map_multi_array_unequal_and_longest
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

$z = array_map(
    fn($a, $b, $c) => "$a-$b-$c",
    [1, 2, 3, 4],
    ['a', 'b'],
    [true, false, true]
);
echo implode('|', $z);

__vybe_check(ob_get_clean(), "1-a-1|2-b-|3--1|4--");
