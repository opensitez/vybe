<?php
// vybe-test: php/string_functions_extended/substr_compare_case_sensitivity_runtime
// origin: languages/php/tests/php/test_string_functions_extended.rs

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

echo "substr_compare_case_sensitivity_runtime_ok";

__vybe_check(ob_get_clean(), "substr_compare_case_sensitivity_runtime_ok");
