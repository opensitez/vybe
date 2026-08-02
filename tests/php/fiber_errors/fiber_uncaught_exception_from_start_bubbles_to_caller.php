<?php
// vybe-test: php/fiber_errors/fiber_uncaught_exception_from_start_bubbles_to_caller
// origin: languages/php/tests/php/test_fiber_errors.rs

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

$f = new Fiber(function (): void { throw new RuntimeException('start-boom'); });
try { $f->start(); echo 'ok'; }
catch (RuntimeException $e) { echo $e->getMessage(); }

__vybe_check(ob_get_clean(), "start-boom");
