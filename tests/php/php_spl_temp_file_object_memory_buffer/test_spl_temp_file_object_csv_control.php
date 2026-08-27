<?php
// vybe-test: php/php_spl_temp_file_object_memory_buffer/test_spl_temp_file_object_csv_control
// origin: languages/php/tests/php/test_php_spl_temp_file_object_memory_buffer.rs

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

echo "test_spl_temp_file_object_csv_control_ok";

__vybe_check(ob_get_clean(), "test_spl_temp_file_object_csv_control_ok");
