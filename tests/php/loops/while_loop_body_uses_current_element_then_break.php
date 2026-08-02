<?php
// vybe-test: php/loops/while_loop_body_uses_current_element_then_break
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
$sum = 0;
$i = 0;
while ($i < count($nums)) {
    $sum += $nums[$i];
    if ($i === 1) { break; }
    $i++;
}
echo $sum;

__vybe_check(ob_get_clean(), "3");
