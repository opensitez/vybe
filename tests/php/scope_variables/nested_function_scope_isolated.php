<?php
// vybe-test: php/scope_variables/nested_function_scope_isolated
// origin: languages/php/tests/php/test_scope_variables.rs

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

function outer(): int {
    $x = 1;
    $inner = function (): int { $y = 2; return $y; };
    return $inner();
}
echo outer();

__vybe_check(ob_get_clean(), "2");
