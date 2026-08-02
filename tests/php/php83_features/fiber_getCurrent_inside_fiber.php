<?php
// vybe-test: php/php83_features/fiber_getCurrent_inside_fiber
// origin: languages/php/tests/php/test_php83_features.rs

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
    echo Fiber::getCurrent() !== null ? 'inside' : 'outside';
    Fiber::suspend();
});
$fiber->start();

__vybe_check(ob_get_clean(), "inside");
