<?php
// vybe-test: php/spl_data_structures/splstack_push_and_pop
// origin: languages/php/tests/php/test_spl_data_structures.rs

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

$s = new SplStack;
$s->push(1); $s->push(2); $s->push(3);
echo $s->pop() . ',' . $s->pop() . ',' . $s->pop();

__vybe_check(ob_get_clean(), "3,2,1");
