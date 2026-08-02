<?php
// vybe-test: php/php_generators_yield_from_send_return/test_php_generator_yield_from_delegation
// origin: languages/php/tests/php/test_php_generators_yield_from_send_return.rs

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

function innerGen() {
    yield 2;
    yield 3;
}

function outerGen() {
    yield 1;
    yield from innerGen();
    yield 4;
}

echo implode(",", iterator_to_array(outerGen()));

__vybe_check(ob_get_clean(), "1,2,3,4");
