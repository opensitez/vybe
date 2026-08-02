<?php
// vybe-test: php/loops/foreach_with_array_shift_side_effect_condition
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

$q = [1, 2, 3, 4];
$sum = 0;
while ($q) {
    $front = array_shift($q);
    if ($front === 2) { continue; }
    $sum += $front;
}
echo $sum;

__vybe_check(ob_get_clean(), "8");
