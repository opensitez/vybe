<?php
// vybe-test: php/generators_send_throw/generator_valid_returns_false_after_exhaustion
// origin: languages/php/tests/php/test_generators_send_throw.rs

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

function two(): Generator { yield 1; yield 2; }
$gen = two();
$gen->current();
$gen->next();
$gen->next();
echo $gen->valid() ? 'yes' : 'no';

__vybe_check(ob_get_clean(), "no");
