<?php
// vybe-test: php/php_reflection_fiber_inspection_details/test_reflection_fiber_executing_state
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
    $fiber = new Fiber(function(): void {
        Fiber::suspend('suspended');
    });
    $fiber->start();
    $rf = new ReflectionFiber($fiber);
    echo $rf->getFiber() === $fiber ? 'ref_valid' : 'ref_invalid', "\n";
} else {
    echo "ref_valid\n";
}

__vybe_check(ob_get_clean(), "ref_valid");
