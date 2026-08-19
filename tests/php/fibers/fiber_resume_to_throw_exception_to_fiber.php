<?php
// vybe-test: php/fibers/fiber_resume_to_throw_exception_to_fiber
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
    try {
        Fiber::suspend('enter');
    } catch (Throwable $e) {
        echo 'caught:' . $e->getMessage();
    }
});
echo $f->start();
try {
    $f->throw(new RuntimeException('from-caller'));
} catch (FiberError) {
    echo '|throw-failed';
}

__vybe_check(ob_get_clean(), "entercaught:from-caller");
