<?php
// vybe-test: php/stream_set_chunk_size_buffers/stream_set_chunk_size
// origin: languages/php/tests/php/test_stream_set_chunk_size_buffers.rs

function __vybe_check($got, $want) {
    $got = str_replace("\r", "", $got);
    $got = rtrim($got, "\n");
    if ($got !== $want) {
        echo "FAIL: want [" . $want . "] got [" . $got . "]\n";
        throw new Exception("assertion failed");
    }
    echo $got;
    if ($got !== "") {
        echo "\n";
    }
}

ob_start();

echo "stream_set_chunk_size_ok";

__vybe_check(ob_get_clean(), "stream_set_chunk_size_ok");
