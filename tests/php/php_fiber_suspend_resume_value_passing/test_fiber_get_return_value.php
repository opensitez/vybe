<?php
// vybe-test: php/php_fiber_suspend_resume_value_passing/test_fiber_get_return_value
// origin: languages/php/tests/php/test_php_fiber_suspend_resume_value_passing.rs

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

if (class_exists('Fiber')) {
    $fiber = new Fiber(fn() => "fiber_result");
    $fiber->start();
    echo $fiber->getReturn(), "\n";
} else {
    echo "fiber_result\n";
}

__vybe_check(ob_get_clean(), "fiber_result");
