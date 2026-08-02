<?php
// vybe-test: php/fibers/fiber_throw_from_within_calls_caller_exception
// origin: languages/php/tests/php/test_fibers.rs

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

$f = new Fiber(function (): void {
    throw new Exception('inside');
});
try {
    $f->start();
    echo 'unreached';
} catch (Exception $e) {
    echo $e->getMessage();
}

__vybe_check(ob_get_clean(), "inside");
