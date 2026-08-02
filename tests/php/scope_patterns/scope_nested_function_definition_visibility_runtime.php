<?php
// vybe-test: php/scope_patterns/scope_nested_function_definition_visibility_runtime
// origin: languages/php/tests/php/test_scope_patterns.rs

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

function define_inner(): void {
    $flag = true;
    if ($flag) {
        function nested_scoped_fn(): string { return 'ok'; }
    }
}
define_inner();
echo function_exists('nested_scoped_fn') ? nested_scoped_fn() : 'missing';

__vybe_check(ob_get_clean(), "ok");
