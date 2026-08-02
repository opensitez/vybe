<?php
// vybe-test: php/error_handler/handler_runs_before_finally_on_trigger_in_try
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
set_error_handler(function() use (&$log): bool { $log[] = 'h'; return true; });
try {
    trigger_error('t', E_USER_NOTICE);
    $log[] = 't';
} finally {
    $log[] = 'f';
}
restore_error_handler();
echo implode('', $log);

__vybe_check(ob_get_clean(), "htf");
