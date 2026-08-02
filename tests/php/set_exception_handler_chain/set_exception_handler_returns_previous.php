<?php
// vybe-test: php/set_exception_handler_chain/set_exception_handler_returns_previous
// origin: languages/php/tests/php/test_set_exception_handler_chain.rs

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

set_exception_handler(function($e) { echo "A"; });
$old = set_exception_handler(function($e) { echo "B"; });

echo is_callable($old) ? "callable" : "not";

__vybe_check(ob_get_clean(), "callable");
