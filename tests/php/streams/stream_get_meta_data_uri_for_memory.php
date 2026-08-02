<?php
// vybe-test: php/streams/stream_get_meta_data_uri_for_memory
// origin: languages/php/tests/php/test_streams.rs

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

$fp = fopen('php://memory', 'r+');
$m = stream_get_meta_data($fp);
echo str_starts_with($m['uri'], 'php://') ? 'php' : 'other';

__vybe_check(ob_get_clean(), "php");
