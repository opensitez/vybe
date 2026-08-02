<?php
// vybe-test: php/loops/while_loop_with_assignment_condition_and_skip
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

$values = [1, 2, 3];
$i = 0;
$sum = 0;
while (($value = $values[$i] ?? null) !== null) {
    $i++;
    if ($value === 2) { continue; }
    $sum += $value;
}
echo $sum;

__vybe_check(ob_get_clean(), "4");
