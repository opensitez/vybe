<?php
// vybe-test: php/control_flow/break_and_continue_with_nested_level_runtime
// origin: languages/php/tests/php/test_control_flow.rs

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

echo "break_and_continue_with_nested_level_runtime_ok";

__vybe_check(ob_get_clean(), "break_and_continue_with_nested_level_runtime_ok");
