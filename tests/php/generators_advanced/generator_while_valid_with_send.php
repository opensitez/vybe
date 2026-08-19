<?php
// vybe-test: php/generators_advanced/generator_while_valid_with_send
// origin: languages/php/tests/php/test_generators_advanced.rs

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

function multiplier() {
    $factor = yield "ready";
    while (true) {
        $value = yield;
        if ($value === null) return;
        yield $value * $factor;
    }
}
$g = multiplier();
$g->current(); // "ready"
$g->send(3);   // sets factor, yields null
$g->next();    // advance to inner yield
echo $g->send(5);   // yields 15
$g->next();
echo $g->send(7);   // yields 21

__vybe_check(ob_get_clean(), "");
