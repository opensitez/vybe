<?php
// vybe-test: php/loops/while_break_2_skips_after_partial_inner_iteration
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
$i = 0;
while (true) {
    $i++;
    foreach ([1, 2] as $v) {
        if ($i === 3 && $v === 2) {
            break 2;
        }
        $sum += $v;
    }
}
echo $sum . '|' . $i;

__vybe_check(ob_get_clean(), "7|3");
