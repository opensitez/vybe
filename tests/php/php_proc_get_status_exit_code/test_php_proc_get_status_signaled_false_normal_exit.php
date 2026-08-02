<?php
// vybe-test: php/php_proc_get_status_exit_code/test_php_proc_get_status_signaled_false_normal_exit
// origin: languages/php/tests/php/test_php_proc_get_status_exit_code.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo hello", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[1]);
    usleep(10000);
    $status = proc_get_status($process);
    proc_close($process);
    echo !$status["signaled"] ? "SIGNALED_FALSE_OK" : "FAIL";
} else {
    echo "SIGNALED_FALSE_OK";
}
