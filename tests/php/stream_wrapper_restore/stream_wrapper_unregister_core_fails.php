<?php
// vybe-test: php/stream_wrapper_restore/stream_wrapper_unregister_core_fails
// origin: languages/php/tests/php/test_stream_wrapper_restore.rs

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

// Attempt to unregister core wrapper should fail or be restored
try {
    $result = stream_wrapper_unregister("http");
    echo $result ? "unregistered" : "failed";
} catch (\Throwable $e) {
    echo "failed";
}

__vybe_check(ob_get_clean(), "unregistered");
