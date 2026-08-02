<?php
// vybe-test: php/first_class_callables/first_class_callable_from_array_map
// origin: languages/php/tests/php/test_first_class_callables.rs

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

$numbers = range(1, 5);
$squared = array_map(fn($n) => $n ** 2, $numbers);
$toStr = strval(...);
$strings = array_map($toStr, $squared);
echo implode(',', $strings) . "\n";

__vybe_check(ob_get_clean(), "1,4,9,16,25");
