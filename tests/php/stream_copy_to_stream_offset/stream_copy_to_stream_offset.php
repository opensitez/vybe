<?php
// vybe-test: php/stream_copy_to_stream_offset/stream_copy_to_stream_offset
// origin: languages/php/tests/php/test_stream_copy_to_stream_offset.rs

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

$src = fopen("php://memory", "w+");
fwrite($src, "Hello, world!");

$dest = fopen("php://memory", "w+");
stream_copy_to_stream($src, $dest, 5, 7);

rewind($dest);
echo stream_get_contents($dest);

__vybe_check(ob_get_clean(), "world");
