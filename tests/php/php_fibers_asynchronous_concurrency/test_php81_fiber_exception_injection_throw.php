<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_exception_injection_throw
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

$fiber = new Fiber(function(): void {
    try {
        Fiber::suspend();
    } catch (RuntimeException $e) {
        echo "CAUGHT_IN_FIBER: " . $e->getMessage();
    }
});

$fiber->start();
$fiber->throw(new RuntimeException("Fiber Exception"));

__vybe_check(ob_get_clean(), "CAUGHT_IN_FIBER: Fiber Exception");
