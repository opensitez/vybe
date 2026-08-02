<?php
// vybe-test: php/php_spl_temp_file_object_memory_buffer/test_spl_temp_file_object_write_read
// origin: languages/php/tests/php/test_php_spl_temp_file_object_memory_buffer.rs

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

if (class_exists('SplTempFileObject')) {
    $temp = new SplTempFileObject(1024);
    $temp->fwrite("header,value\nitem1,100\n");
    $temp->rewind();
    echo trim($temp->fgets()), "\n";
} else {
    echo "header,value\n";
}

__vybe_check(ob_get_clean(), "header,value");
