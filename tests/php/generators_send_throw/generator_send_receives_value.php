<?php
// vybe-test: php/generators_send_throw/generator_send_receives_value
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

function accumulator(): Generator {
    $total = 0;
    while (true) {
        $val = yield $total;
        if ($val === null) break;
        $total += $val;
    }
}
$gen = accumulator();
$gen->current();
$gen->send(10);
$gen->send(20);
echo $gen->send(5);

__vybe_check(ob_get_clean(), "35");
