<?php
// vybe-test: php/error_handler/division_by_zero_triggers_error_or_exception
// origin: languages/php/tests/php/test_error_handler.rs

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

$tag = '';
set_error_handler(function() use (&$tag): bool { $tag = 'handler'; return true; });
try { $x = 1 / 0; echo 'inf'; }
catch (DivisionByZeroError $e) { $tag = 'exception'; }
restore_error_handler();
echo $tag;

__vybe_check(ob_get_clean(), "exception");
