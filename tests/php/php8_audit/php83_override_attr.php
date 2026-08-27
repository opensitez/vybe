<?php
// vybe-test: php/php8_audit/php83_override_attr
// origin: languages/php/tests/php/test_php8_audit.rs

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

echo "php83_override_attr_ok";

__vybe_check(ob_get_clean(), "php83_override_attr_ok");
