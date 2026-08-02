<?php
// vybe-test: php/spl/spl_stack_pop_and_count_runtime
// origin: languages/php/tests/php/test_spl.rs

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
$stack->push('a');
$stack->push('b');
$stack->push('c');
echo $stack->count();
echo '|';
echo $stack->pop();
echo '|';
echo $stack->count();

__vybe_check(ob_get_clean(), "3|c|2");
