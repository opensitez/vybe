<?php
// vybe-test: php/stream_select_read_write/stream_select_write_ready
// origin: languages/php/tests/php/test_stream_select_read_write.rs

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

$fp = fopen("php://temp", "w+");
$read = null;
$write = [$fp];
$except = null;
$num = stream_select($read, $write, $except, 0);
echo $num;

__vybe_check(ob_get_clean(), "1");
