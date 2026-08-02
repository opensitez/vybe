<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_basic_suspend_and_resume
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
    $value = Fiber::suspend("fiber_yield_1");
    echo "received: $value";
});

$res1 = $fiber->start();
echo "start=$res1 | ";
$fiber->resume("resume_val");

__vybe_check(ob_get_clean(), "start=fiber_yield_1 | received: resume_val");
