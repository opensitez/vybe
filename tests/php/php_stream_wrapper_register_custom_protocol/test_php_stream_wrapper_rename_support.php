<?php
// vybe-test: php/php_stream_wrapper_register_custom_protocol/test_php_stream_wrapper_rename_support
// origin: languages/php/tests/php/test_php_stream_wrapper_register_custom_protocol.rs

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

renameproto://a

__vybe_check(ob_get_clean(), "renameproto://b");
