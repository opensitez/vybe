<?php
// vybe-test: php/fiber_errors/fiber_is_terminated_after_uncaught_inner_exception
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

$f = new Fiber(function (): void { throw new Exception('die'); });
try { $f->start(); } catch (Exception $e) { /* handled */ }
echo $f->isTerminated() ? 'dead' : 'alive';

__vybe_check(ob_get_clean(), "dead");
