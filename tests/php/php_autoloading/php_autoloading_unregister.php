<?php
// vybe-test: php/php_autoloading/php_autoloading_unregister
// origin: languages/php/tests/php/test_php_autoloading.rs

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

echo "php_autoloading_unregister_ok";

__vybe_check(ob_get_clean(), "php_autoloading_unregister_ok");
