<?php
// vybe-test: php/list_destructuring/foreach_nested_list_skip_first
// origin: languages/php/tests/php/test_list_destructuring.rs

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

$pairs = [[1,2,3],[4,5,6]];
foreach ($pairs as [,$b,$c]) echo $b . $c;

__vybe_check(ob_get_clean(), "2356");
