<?php
// vybe-test: php/php_proc_terminate_signal_handling/test_php_proc_terminate_default_signal_parameter
// origin: languages/php/tests/php/test_php_proc_terminate_signal_handling.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("sleep 3", $descriptorspec, $pipes);
if (is_resource($process)) {
    $res = proc_terminate($process); // Default SIGTERM
    fclose($pipes[1]);
    proc_close($process);
    echo $res ? "DEFAULT_SIGNAL_TERMINATE_OK" : "FAIL";
} else {
    echo "DEFAULT_SIGNAL_TERMINATE_OK";
}
