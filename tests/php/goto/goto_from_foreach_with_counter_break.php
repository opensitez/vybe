<?php
// vybe-test: php/goto/goto_from_foreach_with_counter_break
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
$i = 0;
foreach ([1, 2, 3] as $v) {
    if ($v === 2) {
        $sum += 5;
        goto after_loop;
    }
    $sum += $v;
}
after_loop:
echo $sum;

__vybe_check(ob_get_clean(), "6");
