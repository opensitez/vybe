<?php
// vybe-test: php/loops/while_assign_in_condition_reads_counter
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

$items = [10, 20];
$i = 0;
$sum = 0;
while (($v = $items[$i] ?? null) !== null) {
    $sum += $v;
    $i++;
}
echo $sum;

__vybe_check(ob_get_clean(), "30");
