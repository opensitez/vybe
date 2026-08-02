<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_terminate_closed_process_returns_false
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo done", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[1]);
    proc_close($process);
    $res = @proc_terminate($process);
    echo $res === false ? "CLOSED_TERMINATE_FALSE_OK" : "FAIL";
} else {
    echo "CLOSED_TERMINATE_FALSE_OK";
}
