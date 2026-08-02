<?php
// vybe-test: php/loops/for_loop_with_continue_increments_before_skip
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

$sum = 0;
for ($i = 0; $i < 6; $i++) {
    $sum += $i;
    if ($i % 2 === 0) {
        continue;
    }
    $sum += 10;
}
echo $sum;

__vybe_check(ob_get_clean(), "45");
