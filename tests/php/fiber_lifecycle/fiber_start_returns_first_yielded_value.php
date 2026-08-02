<?php
// vybe-test: php/fiber_lifecycle/fiber_start_returns_first_yielded_value
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
    $val = Fiber::suspend('first');
    echo "got: $val";
});
$result = $fiber->start();
echo $result;

__vybe_check(ob_get_clean(), "first");
