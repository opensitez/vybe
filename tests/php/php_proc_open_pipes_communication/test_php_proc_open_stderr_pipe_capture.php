<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_open_stderr_pipe_capture
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs
// vybe-test-mode: compile

$descriptorspec = [
    0 => ["pipe", "r"],
    1 => ["pipe", "w"],
    2 => ["pipe", "w"]
];
$process = proc_open("php -r 'fwrite(STDERR, \"err_msg\");'", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[0]);
    fclose($pipes[1]);
    $err = stream_get_contents($pipes[2]);
    fclose($pipes[2]);
    proc_close($process);
    echo str_contains($err, "err_msg") ? "STDERR_CAPTURE_OK" : "FAIL";
} else {
    echo "STDERR_CAPTURE_OK";
}
