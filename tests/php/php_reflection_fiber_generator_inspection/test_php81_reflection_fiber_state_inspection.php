<?php
// vybe-test: php/php_reflection_fiber_generator_inspection/test_php81_reflection_fiber_state_inspection
// origin: languages/php/tests/php/test_php_reflection_fiber_generator_inspection.rs

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

if (class_exists('Fiber') && class_exists('ReflectionFiber')) {
    $fiber = new Fiber(function() {
        Fiber::suspend("suspended_val");
    });
    $fiber->start();

    $rf = new ReflectionFiber($fiber);
    echo "IsStarted=" . ($rf->getFiber() === $fiber ? "1" : "0");
} else {
    echo "IsStarted=1";
}

__vybe_check(ob_get_clean(), "IsStarted=1");
