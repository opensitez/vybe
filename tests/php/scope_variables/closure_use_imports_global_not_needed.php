<?php
// vybe-test: php/scope_variables/closure_use_imports_global_not_needed
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

$outer = 3;
$fn = function () use ($outer): int { return $outer; };
echo $fn();

__vybe_check(ob_get_clean(), "3");
