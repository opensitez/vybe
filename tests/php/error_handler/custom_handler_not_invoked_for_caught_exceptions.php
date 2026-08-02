<?php
// vybe-test: php/error_handler/custom_handler_not_invoked_for_caught_exceptions
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

$hit = false;
set_error_handler(function() use (&$hit): bool { $hit = true; return true; });
try { throw new RuntimeException('ex'); }
catch (RuntimeException $e) { echo 'catch'; }
restore_error_handler();
echo $hit ? ':handler' : ':clean';

__vybe_check(ob_get_clean(), "catch:clean");
