<?php
// vybe-test: php/php_proc_get_status_exit_code/test_php_proc_get_status_stopped_false_normal_run
// origin: languages/php/tests/php/test_php_proc_get_status_exit_code.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo hello", $descriptorspec, $pipes);
if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo !$status["stopped"] ? "STOPPED_FALSE_OK" : "FAIL";
} else {
    echo "STOPPED_FALSE_OK";
}
