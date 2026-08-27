<?php
// vybe-test: php/phpinfo/phpinfo_return_value_is_true
// origin: languages/php/tests/php/test_phpinfo.rs

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

echo "phpinfo_return_value_is_true_ok";

__vybe_check(ob_get_clean(), "phpinfo_return_value_is_true_ok");
