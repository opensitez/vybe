<?php
// vybe-test: php/error_handler/ini_set_display_errors_does_not_break_handler
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

$old = ini_set('display_errors', '0');
$hit = false;
set_error_handler(function() use (&$hit): bool { $hit = true; return true; });
trigger_error('quiet', E_USER_NOTICE);
restore_error_handler();
ini_set('display_errors', $old !== false ? $old : '1');
echo $hit ? 'hit' : 'miss';

__vybe_check(ob_get_clean(), "hit");
