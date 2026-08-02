<?php
// vybe-test: php/operators_runtime/mixed_assignment_and_short_circuit_runtime
// origin: languages/php/tests/php/test_operators_runtime.rs

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

$count = 0;
$count = $count ?: 1;
echo $count . '|';
$count = 0;
echo ($count && ($count = 9)) . '|';
echo $count . '|';
$count = 1;
echo ($count && ($count = 9)) . '|';
echo $count;

__vybe_check(ob_get_clean(), "1||0|1|9");
