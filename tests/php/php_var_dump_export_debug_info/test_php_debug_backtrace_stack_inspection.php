<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_debug_backtrace_stack_inspection
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs

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

function foo($a, $b) { bar($a + $b); }
function bar($c) { baz($c); }
function baz($d) {
    $trace = debug_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS);
    $funcs = array_column($trace, "function");
    echo implode(" <- ", $funcs);
}

foo(10, 20);

__vybe_check(ob_get_clean(), "baz <- bar <- foo");
