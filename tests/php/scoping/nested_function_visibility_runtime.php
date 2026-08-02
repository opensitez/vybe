<?php
// vybe-test: php/scoping/nested_function_visibility_runtime
// origin: languages/php/tests/php/test_scoping.rs

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

$prefix = 'outer';
function outer_scope(): string {
    function inner_scope_fn(): string { return 'inner'; }
    return inner_scope_fn();
}
echo outer_scope() . '|' . inner_scope_fn();

__vybe_check(ob_get_clean(), "inner|inner");
