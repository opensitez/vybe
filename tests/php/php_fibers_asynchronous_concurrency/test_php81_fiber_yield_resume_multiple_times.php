<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_yield_resume_multiple_times
// origin: languages/php/tests/php/test_php_fibers_asynchronous_concurrency.rs

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

$sequence = [];
$fiber = new Fiber(function() use (&$sequence) {
    $sequence[] = 'start';
    $v = Fiber::suspend('first');
    $sequence[] = $v;
    $v = Fiber::suspend('second');
    $sequence[] = $v;
    return 'done';
});
$sequence[] = $fiber->start();
$fiber->resume('r1');
$fiber->resume('r2');
$sequence[] = $fiber->getReturn();
echo implode('|', $sequence);

__vybe_check(ob_get_clean(), "start|first|r1|second|r2|done");
