<?php
// vybe-test: php/loops/foreach_with_nested_list_destroys_reference_after_loop
// origin: languages/php/tests/php/test_loops.rs

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

$items = [[1], [2], [3]];
$sum = 0;
foreach ($items as &$pair) {
    $pair[0] *= 2;
}
unset($pair);
foreach ($items as $pair) {
    $sum += $pair[0];
}
echo $sum;

__vybe_check(ob_get_clean(), "7");
