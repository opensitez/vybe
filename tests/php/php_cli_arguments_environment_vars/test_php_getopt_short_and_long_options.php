<?php
// vybe-test: php/php_cli_arguments_environment_vars/test_php_getopt_short_and_long_options
// origin: languages/php/tests/php/test_php_cli_arguments_environment_vars.rs

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

echo "test_php_getopt_short_and_long_options_ok";

__vybe_check(ob_get_clean(), "test_php_getopt_short_and_long_options_ok");
