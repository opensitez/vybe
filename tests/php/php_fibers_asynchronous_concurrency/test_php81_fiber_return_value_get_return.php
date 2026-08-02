<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_return_value_get_return
// origin: languages/php/tests/php/test_php_fibers_asynchronous_concurrency.rs

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

$fiber = new Fiber(fn() => "fiber_result");
$fiber->start();
echo $fiber->getReturn();

__vybe_check(ob_get_clean(), "fiber_result");
