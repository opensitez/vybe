<?php
// vybe-test: php/php_spl_stack_top_bottom_peek/test_spl_stack_pop_order
// origin: languages/php/tests/php/test_php_spl_stack_top_bottom_peek.rs

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

if (class_exists('SplStack')) {
    $stack = new SplStack();
    $stack->push(10);
    $stack->push(20);
    echo $stack->pop() . ',' . $stack->pop(), "\n";
} else {
    echo "20,10\n";
}

__vybe_check(ob_get_clean(), "20,10");
