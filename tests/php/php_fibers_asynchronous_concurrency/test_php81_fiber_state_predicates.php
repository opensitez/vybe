<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_state_predicates
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

$fiber = new Fiber(function(): int {
    Fiber::suspend();
    return 42;
});

echo $fiber->isStarted() ? "0" : "1";
$fiber->start();
echo $fiber->isSuspended() ? "1" : "0";
$res = $fiber->resume();
echo $fiber->isTerminated() ? "1" : "0";
echo " res=$res";

__vybe_check(ob_get_clean(), "111 res=42");
