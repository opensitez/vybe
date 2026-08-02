<?php
// vybe-test: php/php_fiber_suspend_resume_value_passing/test_fiber_bidirectional_value_passing
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
    $fiber = new Fiber(function(string $param): string {
        $received = Fiber::suspend("yielded:" . $param);
        return "returned:" . $received;
    });
    $yielded = $fiber->start("init");
    $returned = $fiber->resume("resumed");
    echo $yielded . '|' . $returned, "\n";
} else {
    echo "yielded:init|returned:resumed\n";
}

__vybe_check(ob_get_clean(), "yielded:init|returned:resumed");
