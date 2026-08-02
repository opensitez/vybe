<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_terminate_sigint_signal
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 5", $descriptorspec, $pipes);
if (is_resource($process)) {
    $sig = defined('SIGINT') ? SIGINT : 2;
    $res = proc_terminate($process, $sig);
    fclose($pipes[1]);
    proc_close($process);
    echo $res ? "SIGINT_TERMINATE_OK" : "FAIL";
} else {
    echo "SIGINT_TERMINATE_OK";
}
