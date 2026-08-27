<?php
// vybe-test: php/builtins/wordpress_php_version_error_branch_runtime
// origin: languages/php/tests/php/test_builtins.rs

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

echo "wordpress_php_version_error_branch_runtime_ok";

__vybe_check(ob_get_clean(), "wordpress_php_version_error_branch_runtime_ok");
