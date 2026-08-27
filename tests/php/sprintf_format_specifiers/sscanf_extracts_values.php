<?php
// vybe-test: php/sprintf_format_specifiers/sscanf_extracts_values
// origin: languages/php/tests/php/test_sprintf_format_specifiers.rs

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

echo "sscanf_extracts_values_ok";

__vybe_check(ob_get_clean(), "sscanf_extracts_values_ok");
