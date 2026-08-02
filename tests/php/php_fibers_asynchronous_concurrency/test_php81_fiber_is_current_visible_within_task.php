<?php
// vybe-test: php/php_fibers_asynchronous_concurrency/test_php81_fiber_is_current_visible_within_task
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

$label = '';
$fiber = new Fiber(function() use (&$label): void {
    $current = Fiber::getCurrent();
    if ($current !== null) {
        $label = 'has_current';
    } else {
        $label = 'none';
    }
});
$fiber->start();
echo $label;

__vybe_check(ob_get_clean(), "has_current");
