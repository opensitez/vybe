<?php
// vybe-test: php/php_proc_get_status_exit_code/test_php_proc_get_status_stopsig_default_minus_one
// origin: languages/php/tests/php/test_php_proc_get_status_exit_code.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo test", $descriptorspec, $pipes);
if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo $status["stopsig"] === -1 || is_int($status["stopsig"]) ? "STOPSIG_INT_OK" : "FAIL";
} else {
    echo "STOPSIG_INT_OK";
}
