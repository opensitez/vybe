<?php
// vybe-test: php/php_output_buffering_nested_handlers/test_php_ob_start_flags_erase_write_flush
// origin: languages/php/tests/php/test_php_output_buffering_nested_handlers.rs

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

echo "test_php_ob_start_flags_erase_write_flush_ok";

__vybe_check(ob_get_clean(), "test_php_ob_start_flags_erase_write_flush_ok");
