<?php
// vybe-test: php/loops/foreach_reference_then_value_iterates_original_values
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

$nums = [1, 2, 3];
foreach ($nums as &$n) {
    $n *= 2;
}
unset($n);
$sum = 0;
foreach ($nums as $n) {
    $sum += $n;
}
echo $sum;

__vybe_check(ob_get_clean(), "12");
