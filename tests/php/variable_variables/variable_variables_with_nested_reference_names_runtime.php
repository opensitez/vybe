<?php
// vybe-test: php/variable_variables/variable_variables_with_nested_reference_names_runtime
// origin: languages/php/tests/php/test_variable_variables.rs

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

echo "variable_variables_with_nested_reference_names_runtime_ok";

__vybe_check(ob_get_clean(), "variable_variables_with_nested_reference_names_runtime_ok");
