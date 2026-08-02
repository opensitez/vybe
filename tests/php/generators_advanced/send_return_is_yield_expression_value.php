<?php
// vybe-test: php/generators_advanced/send_return_is_yield_expression_value
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

function doubler() {
    while (true) {
        $input = yield;
        if ($input === null) return;
        yield $input * 2;
    }
}
$g = doubler();
$g->current(); // prime
echo $g->send(5);  // yields 10
$g->next();        // advance to next yield
echo $g->send(7);  // yields 14

__vybe_check(ob_get_clean(), "1014");
