<?php
// vybe-test: php/error_handler/trigger_in_catch_block_still_handled
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

$log = [];
set_error_handler(function() use (&$log): bool { $log[] = 'e'; return true; });
try { throw new Exception('x'); }
catch (Exception $ex) {
    trigger_error('in-catch', E_USER_NOTICE);
    $log[] = 'c';
}
restore_error_handler();
echo implode('', $log);

__vybe_check(ob_get_clean(), "ec");
