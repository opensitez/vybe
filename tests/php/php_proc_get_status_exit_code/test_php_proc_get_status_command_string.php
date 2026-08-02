<?php
// vybe-test: php/php_proc_get_status_exit_code/test_php_proc_get_status_command_string
// origin: languages/php/tests/php/test_php_proc_get_status_exit_code.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo 'test_command_prop'", $descriptorspec, $pipes);
if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo str_contains($status["command"], "test_command_prop") ? "COMMAND_PROP_OK" : "FAIL";
} else {
    echo "COMMAND_PROP_OK";
}
