<?php
// vybe-test: php/goto/goto_exits_loop_early
// origin: languages/php/tests/php/test_goto.rs

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
for ($i = 1; $i <= 10; $i++) {
    if ($i > 3) goto stop;
    $sum += $i;
}
stop:
echo $sum;

__vybe_check(ob_get_clean(), "6");
