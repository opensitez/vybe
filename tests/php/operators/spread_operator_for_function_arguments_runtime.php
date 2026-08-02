<?php
// vybe-test: php/operators/spread_operator_for_function_arguments_runtime
// origin: languages/php/tests/php/test_operators.rs

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

function sum(int ...$nums): int {
    $total = 0;
    foreach ($nums as $n) { $total += $n; }
    return $total;
}
$numbers = [1, 2, 3];
echo sum(...$numbers);

__vybe_check(ob_get_clean(), "6");
