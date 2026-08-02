<?php
// vybe-test: php/named_arguments/named_arg_with_array_value
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

function sumArray(array $items, int $initial = 0): int {
    return array_reduce($items, fn($c, $v) => $c + $v, $initial);
}
echo sumArray(items: [1, 2, 3, 4]) . "\n";
echo sumArray(items: [10, 20], initial: 5) . "\n";

__vybe_check(ob_get_clean(), "10\n35");
