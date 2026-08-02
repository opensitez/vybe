<?php
// vybe-test: php/loops/while_loop_with_assignment_in_condition_and_truthiness
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

$values = [3, 2, 0, 5];
$i = 0;
$sum = '';
while (($v = $values[$i] ?? null)) {
    $sum .= $v;
    $i++;
}
echo $sum;

__vybe_check(ob_get_clean(), "32");
