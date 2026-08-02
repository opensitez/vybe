<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_open_blocking_pipes
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs
// vybe-test-mode: compile

$descriptorspec = [0 => ["pipe", "r"], 1 => ["pipe", "w"]];
$process = proc_open("cat", $descriptorspec, $pipes);
if (is_resource($process)) {
    stream_set_blocking($pipes[1], false);
    fclose($pipes[0]);
    fclose($pipes[1]);
    proc_close($process);
    echo "NON_BLOCKING_PIPE_OK";
} else {
    echo "NON_BLOCKING_PIPE_OK";
}
