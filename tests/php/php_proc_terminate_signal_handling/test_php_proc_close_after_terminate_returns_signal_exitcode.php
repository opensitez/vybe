<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_close_after_terminate_returns_signal_exitcode
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 5", $descriptorspec, $pipes);
if (is_resource($process)) {
    proc_terminate($process, 9);
    fclose($pipes[1]);
    $code = proc_close($process);
    echo is_int($code) ? "TERMINATE_CLOSE_CODE_OK" : "FAIL";
} else {
    echo "TERMINATE_CLOSE_CODE_OK";
}
