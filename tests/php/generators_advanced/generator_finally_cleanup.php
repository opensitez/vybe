<?php
// vybe-test: php/generators_advanced/generator_finally_cleanup
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

function resource() {
    echo "open";
    try {
        yield 1;
        yield 2;
    } finally {
        echo "close";
    }
}
$gen = resource();
echo $gen->current();
$gen->next();
echo $gen->current();
$gen = null; // drop triggers finally

__vybe_check(ob_get_clean(), "open12close");
