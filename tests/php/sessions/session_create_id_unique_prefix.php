<?php
// vybe-test: php/sessions/session_create_id_unique_prefix
// origin: languages/php/tests/php/test_sessions.rs

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

$id = session_create_id('app');
echo str_starts_with($id, 'app') ? 'pref' : 'nop';

__vybe_check(ob_get_clean(), "pref");
