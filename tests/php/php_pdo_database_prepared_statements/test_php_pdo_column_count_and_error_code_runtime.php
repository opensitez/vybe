<?php
// vybe-test: php/php_pdo_database_prepared_statements/test_php_pdo_column_count_and_error_code_runtime
// origin: languages/php/tests/php/test_php_pdo_database_prepared_statements.rs

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

echo "test_php_pdo_column_count_and_error_code_runtime_ok";

__vybe_check(ob_get_clean(), "test_php_pdo_column_count_and_error_code_runtime_ok");
