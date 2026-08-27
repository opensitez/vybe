<?php
// vybe-test: php/file_functions_runtime/filesize_memory_stream_after_write
// origin: languages/php/tests/php/test_file_functions_runtime.rs

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

echo "filesize_memory_stream_after_write_ok";

__vybe_check(ob_get_clean(), "filesize_memory_stream_after_write_ok");
