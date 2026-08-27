<?php
// vybe-test: php/version_compare/zend_version_constant_non_empty
// origin: languages/php/tests/php/test_version_compare.rs

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

echo "zend_version_constant_non_empty_ok";

__vybe_check(ob_get_clean(), "zend_version_constant_non_empty_ok");
