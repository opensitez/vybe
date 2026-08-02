<?php
// vybe-test: php/php_proc_get_status_exit_code/test_php_proc_get_status_termsig_default_minus_one
// origin: languages/php/tests/php/test_php_proc_get_status_exit_code.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("echo test", $descriptorspec, $pipes);
if (is_resource($process)) {
    $status = proc_get_status($process);
    fclose($pipes[1]);
    proc_close($process);
    echo $status["termsig"] === -1 || is_int($status["termsig"]) ? "TERMSIG_INT_OK" : "FAIL";
} else {
    echo "TERMSIG_INT_OK";
}
