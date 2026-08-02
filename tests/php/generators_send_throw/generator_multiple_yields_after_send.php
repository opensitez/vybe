<?php
// vybe-test: php/generators_send_throw/generator_multiple_yields_after_send
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

function multiStep(): Generator {
    $a = yield 'first';
    $b = yield 'second';
    yield "$a+$b=" . ($a + $b);
}
$gen = multiStep();
$gen->current();
$gen->send(3);
echo $gen->send(7);

__vybe_check(ob_get_clean(), "3+7=10");
