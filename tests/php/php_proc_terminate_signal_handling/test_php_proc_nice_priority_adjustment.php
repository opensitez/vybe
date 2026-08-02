<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_nice_priority_adjustment
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs
// vybe-test-mode: compile

if (function_exists('proc_nice')) {
    $res = @proc_nice(0);
    echo is_bool($res) ? "PROC_NICE_OK" : "FAIL";
} else {
    echo "PROC_NICE_OK";
}
