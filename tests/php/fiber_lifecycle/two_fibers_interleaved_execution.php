<?php
// vybe-test: php/fiber_lifecycle/two_fibers_interleaved_execution
// origin: languages/php/tests/php/test_fiber_lifecycle.rs

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

$log = [];
$a = new Fiber(function() use (&$log): void {
    $log[] = 'A1'; Fiber::suspend();
    $log[] = 'A2'; Fiber::suspend();
    $log[] = 'A3';
});
$b = new Fiber(function() use (&$log): void {
    $log[] = 'B1'; Fiber::suspend();
    $log[] = 'B2';
});
$a->start(); $b->start();
$a->resume(); $b->resume();
$a->resume();
echo implode(',', $log);

__vybe_check(ob_get_clean(), "A1,B1,A2,B2,A3");
