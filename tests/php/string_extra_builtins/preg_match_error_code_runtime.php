<?php
// vybe-test: php/string_extra_builtins/preg_match_error_code_runtime
// origin: languages/php/tests/php/test_string_extra_builtins.rs

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

echo "preg_match_error_code_runtime_ok";

__vybe_check(ob_get_clean(), "preg_match_error_code_runtime_ok");
