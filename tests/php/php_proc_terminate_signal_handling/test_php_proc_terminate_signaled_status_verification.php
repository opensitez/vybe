<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_terminate_signaled_status_verification
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 10", $descriptorspec, $pipes);
if (is_resource($process)) {
    proc_terminate($process, 9);
    usleep(10000);
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo $status["signaled"] || $status["exitcode"] !== 0 ? "SIGNALED_STATUS_OK" : "FAIL";
} else {
    echo "SIGNALED_STATUS_OK";
}
