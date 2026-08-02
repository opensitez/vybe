<?php
// vybe-test: php/fibers/fiber_get_return_after_exception_is_error
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
    throw new RuntimeException('boom');
});
try {
    $f->start();
} catch (RuntimeException $e) {
    echo $e->getMessage();
}
try {
    $f->getReturn();
    echo '|bad';
} catch (FiberError) {
    echo '|blocked';
}

__vybe_check(ob_get_clean(), "boom|blocked");
