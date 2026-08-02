<?php
// vybe-test: php/php_proc_get_status_exit_code/test_php_proc_get_status_repeated_calls
// origin: languages/php/tests/php/test_php_proc_get_status_exit_code.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 1", $descriptorspec, $pipes);
if (is_resource($process)) {
    $s1 = proc_get_status($process);
    $s2 = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo $s1["pid"] === $s2["pid"] ? "REPEATED_STATUS_PID_OK" : "FAIL";
} else {
    echo "REPEATED_STATUS_PID_OK";
}
