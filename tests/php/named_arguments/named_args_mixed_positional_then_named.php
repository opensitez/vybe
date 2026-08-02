<?php
// vybe-test: php/named_arguments/named_args_mixed_positional_then_named
// origin: languages/php/tests/php/test_named_arguments.rs

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

function rangeSum(int $start, int $end, int $step = 1): int {
    $sum = 0;
    for ($i = $start; $i <= $end; $i += $step) $sum += $i;
    return $sum;
}
echo rangeSum(1, 10) . "\n";
echo rangeSum(1, end: 10, step: 2) . "\n";

__vybe_check(ob_get_clean(), "55\n25");
