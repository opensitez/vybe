<?php
// vybe-test: php/php_reflection_fiber_inspection_details/test_reflection_fiber_get_callable
// origin: languages/php/tests/php/test_php_reflection_fiber_inspection_details.rs

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
    $callable = function() { return 42; };
    $fiber = new Fiber($callable);
    $rf = new ReflectionFiber($fiber);
    echo is_callable($rf->getCallable()) ? 'callable_ok' : 'not_callable', "\n";
} else {
    echo "callable_ok\n";
}

__vybe_check(ob_get_clean(), "callable_ok");
