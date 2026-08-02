<?php
// vybe-test: php/php_proc_get_status_exit_code/test_php_proc_get_status_closed_resource_returns_false
// origin: languages/php/tests/php/test_php_proc_get_status_exit_code.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo done", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[1]);
    proc_close($process);
    $res = @proc_get_status($process);
    echo $res === false ? "CLOSED_STATUS_FALSE_OK" : "FAIL";
} else {
    echo "CLOSED_STATUS_FALSE_OK";
}
