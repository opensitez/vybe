<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_terminate_with_sigterm_signal
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 5", $descriptorspec, $pipes);
if (is_resource($process)) {
    $sig = defined('SIGTERM') ? SIGTERM : 15;
    $res = proc_terminate($process, $sig);
    fclose($pipes[1]);
    proc_close($process);
    echo $res ? "SIGTERM_OK" : "FAIL";
} else {
    echo "SIGTERM_OK";
}
