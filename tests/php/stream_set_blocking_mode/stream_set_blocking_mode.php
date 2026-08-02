<?php
// vybe-test: php/stream_set_blocking_mode/stream_set_blocking_mode
// origin: languages/php/tests/php/test_stream_set_blocking_mode.rs

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

$fp = fopen("php://temp", "r+");
echo stream_set_blocking($fp, false) ? "ok|" : "fail|";
$meta = stream_get_meta_data($fp);
echo $meta['blocked'] ? "blocked" : "unblocked";

__vybe_check(ob_get_clean(), "ok|unblocked");
