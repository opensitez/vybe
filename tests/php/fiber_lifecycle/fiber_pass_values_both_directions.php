<?php
// vybe-test: php/fiber_lifecycle/fiber_pass_values_both_directions
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

$fiber = new Fiber(function(): string {
    $x = Fiber::suspend("need input");
    $y = Fiber::suspend("got: $x");
    return "final: " . ($x + $y);
});
$prompt1 = $fiber->start();
echo $prompt1 . "\n";
$prompt2 = $fiber->resume(10);
echo $prompt2 . "\n";
$fiber->resume(5);
echo $fiber->getReturn();

__vybe_check(ob_get_clean(), "need input\ngot: 10\nfinal: 15");
