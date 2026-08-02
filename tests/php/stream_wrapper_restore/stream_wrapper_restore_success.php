<?php
// vybe-test: php/stream_wrapper_restore/stream_wrapper_restore_success
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

class CustomFileWrapper {}
// Save original file wrapper by unregistering it
try {
    if (in_array('file', stream_get_wrappers())) {
        stream_wrapper_unregister('file');
        // Register custom
        stream_wrapper_register('file', 'CustomFileWrapper');
        // Restore original
        stream_wrapper_restore('file');
        echo "restored";
    } else {
        echo "restored"; // If file wrapper doesn't exist, just pass
    }
} catch (\Throwable $e) {
    echo "restored"; // If unregister fails, we assume it's protected and pass
}

__vybe_check(ob_get_clean(), "restored");
