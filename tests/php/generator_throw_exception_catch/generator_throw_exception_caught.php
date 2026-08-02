<?php
// vybe-test: php/generator_throw_exception_catch/generator_throw_exception_caught
// origin: languages/php/tests/php/test_generator_throw_exception_catch.rs

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

function gen() {
    try {
        yield 1;
    } catch (\Exception $e) {
        yield $e->getMessage();
    }
}
$g = gen();
echo $g->current() . "|";
$g->throw(new \Exception("thrown"));
echo $g->current();

__vybe_check(ob_get_clean(), "1|thrown");
