<?php
// vybe-test: php/phpinfo/php_version_id_is_80000
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

echo "php_version_id_is_80000_ok";

__vybe_check(ob_get_clean(), "php_version_id_is_80000_ok");
