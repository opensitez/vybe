<?php
// vybe-test: php/error_handler/handler_mask_ignores_notice_but_catches_warning
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

$hits = '';
set_error_handler(function(int $no) use (&$hits): bool {
    $hits .= $no === E_USER_WARNING ? 'W' : 'N';
    return true;
}, E_USER_WARNING);
trigger_error('notice', E_USER_NOTICE);
trigger_error('warning', E_USER_WARNING);
restore_error_handler();
echo $hits;

__vybe_check(ob_get_clean(), "W");
