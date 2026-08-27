<?php
// vybe-test: php/ini_functions/get_cfg_var_matches_ini_get_for_php_version
// origin: languages/php/tests/php/test_ini_functions.rs

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

echo "get_cfg_var_matches_ini_get_for_php_version_ok";

__vybe_check(ob_get_clean(), "get_cfg_var_matches_ini_get_for_php_version_ok");
