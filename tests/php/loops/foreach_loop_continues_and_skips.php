<?php
// vybe-test: php/loops/foreach_loop_continues_and_skips
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
foreach ([1, 2, 3, 4] as $v) {
    if ($v % 2 === 0) {
        continue;
    }
    $sum += $v;
}
echo $sum;

__vybe_check(ob_get_clean(), "4");
