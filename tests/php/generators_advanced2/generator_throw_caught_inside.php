<?php
// vybe-test: php/generators_advanced2/generator_throw_caught_inside
// origin: languages/php/tests/php/test_generators_advanced2.rs

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

function resilient(): Generator {
    try {
        yield 1;
    } catch (RuntimeException $e) {
        yield 'caught:' . $e->getMessage();
    }
    yield 2;
}
$g = resilient();
echo $g->current() . ',';
$g->throw(new RuntimeException('boom'));
echo $g->current() . ',';
$g->next();
echo $g->current();

__vybe_check(ob_get_clean(), "1,caught:boom,2");
