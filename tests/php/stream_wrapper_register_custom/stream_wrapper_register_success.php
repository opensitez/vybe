<?php
// vybe-test: php/stream_wrapper_register_custom/stream_wrapper_register_success
// origin: languages/php/tests/php/test_stream_wrapper_register_custom.rs

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

class MyWrapper {
    public function stream_open($path, $mode, $options, &$opened_path) {
        return true;
    }
}
$result = stream_wrapper_register("myproto", "MyWrapper");
echo $result ? "registered" : "failed";

__vybe_check(ob_get_clean(), "registered");
