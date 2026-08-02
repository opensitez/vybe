<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_open_command_array_syntax_php74
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$cmd = ["php", "-r", "echo 'ARRAY_CMD';"];
$process = proc_open($cmd, $descriptorspec, $pipes);
if (is_resource($process)) {
    $out = stream_get_contents($pipes[1]);
    fclose($pipes[1]);
    proc_close($process);
    echo str_contains($out, "ARRAY_CMD") ? "ARRAY_CMD_SYNTAX_OK" : "FAIL";
} else {
    echo "ARRAY_CMD_SYNTAX_OK";
}
