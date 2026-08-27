<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_terminate_closed_process_returns_false
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs

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

echo "test_php_proc_terminate_closed_process_returns_false_ok";

__vybe_check(ob_get_clean(), "test_php_proc_terminate_closed_process_returns_false_ok");
