<?php
// vybe-test: php/fibers/fiber_suspend_then_terminate_flow
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

$f = new Fiber(function (): int {
    echo Fiber::suspend('a');
    Fiber::suspend('b');
    return 42;
});
echo $f->start();
echo '|';
echo $f->resume('x');
echo '|';
echo $f->resume('y');

__vybe_check(ob_get_clean(), "a|xb|");
