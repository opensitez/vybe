<?php
// vybe-test: php/php_connection_status_user_abort_ops/test_ignore_user_abort_toggle
// origin: languages/php/tests/php/test_php_connection_status_user_abort_ops.rs

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

$prev = ignore_user_abort(true);
$now = ignore_user_abort(false);
echo ($now === 1 || $now === true) ? 'toggled' : 'err', "\n";

__vybe_check(ob_get_clean(), "toggled");
