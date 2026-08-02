<?php
// vybe-test: php/fiber_lifecycle/fiber_resume_sends_value_and_gets_next_suspend
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

$fiber = new Fiber(function(): void {
    $a = Fiber::suspend(1);
    $b = Fiber::suspend(2);
    echo "$a,$b";
});
$fiber->start();
$fiber->resume('x');
$fiber->resume('y');

__vybe_check(ob_get_clean(), "x,y");
