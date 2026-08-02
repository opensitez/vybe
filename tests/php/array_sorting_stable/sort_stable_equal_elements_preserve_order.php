<?php
// vybe-test: php/array_sorting_stable/sort_stable_equal_elements_preserve_order
// origin: languages/php/tests/php/test_array_sorting_stable.rs

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

$items = [['n'=>'b','v'=>2],['n'=>'a','v'=>2],['n'=>'c','v'=>1]];
usort($items, fn($a,$b) => $a['v'] <=> $b['v']);
echo $items[0]['n'] . ',' . $items[1]['n'] . ',' . $items[2]['n'];

__vybe_check(ob_get_clean(), "c,b,a");
