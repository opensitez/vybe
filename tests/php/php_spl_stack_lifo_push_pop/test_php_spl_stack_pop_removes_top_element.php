<?php
// vybe-test: php/php_spl_stack_lifo_push_pop/test_php_spl_stack_pop_removes_top_element
// origin: languages/php/tests/php/test_php_spl_stack_lifo_push_pop.rs

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
$stack->push(10);
$stack->push(20);
$val = $stack->pop();
echo "Popped=$val Count=" . $stack->count();

__vybe_check(ob_get_clean(), "Popped=20 Count=1");
