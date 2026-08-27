<?php
// vybe-test: php/spl_autoload/spl_autoload_register_with_array_callback
// origin: languages/php/tests/php/test_spl_autoload.rs

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

echo "spl_autoload_register_with_array_callback_ok";

__vybe_check(ob_get_clean(), "spl_autoload_register_with_array_callback_ok");
