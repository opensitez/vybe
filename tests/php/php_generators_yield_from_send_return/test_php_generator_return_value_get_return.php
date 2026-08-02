<?php
// vybe-test: php/php_generators_yield_from_send_return/test_php_generator_return_value_get_return
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

function countAndReturn() {
    yield 10;
    yield 20;
    return "done_counting";
}

$g = countAndReturn();
foreach ($g as $v) {}
echo $g->getReturn();

__vybe_check(ob_get_clean(), "done_counting");
