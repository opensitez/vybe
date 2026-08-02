<?php
// vybe-test: php/error_handler/handler_with_typed_params_accepts_int_errno
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

$type = '';
set_error_handler(function(int $errno, string $errstr) use (&$type): bool {
    $type = is_int($errno) ? 'int' : 'other';
    return true;
});
trigger_error('typed', E_USER_WARNING);
restore_error_handler();
echo $type;

__vybe_check(ob_get_clean(), "int");
