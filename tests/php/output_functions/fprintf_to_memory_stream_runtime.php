<?php
// vybe-test: php/output_functions/fprintf_to_memory_stream_runtime
// origin: languages/php/tests/php/test_output_functions.rs

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

$stream = fopen('php://memory', 'r+');
fprintf($stream, 'value=%d', 11);
rewind($stream);
$data = stream_get_contents($stream);
echo $data;

__vybe_check(ob_get_clean(), "value=11");
