<?php
// vybe-test: php/php_spl_data_structures_fixed_array/test_php_spl_stack_push_pop_lifo
// origin: languages/php/tests/php/test_php_spl_data_structures_fixed_array.rs

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

$stack = new SplStack();
$stack->push("first");
$stack->push("second");
echo $stack->pop() . " -> " . $stack->pop();

__vybe_check(ob_get_clean(), "second -> first");
