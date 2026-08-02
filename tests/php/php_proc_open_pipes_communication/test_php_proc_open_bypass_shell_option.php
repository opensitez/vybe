<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_open_bypass_shell_option
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$options = ["bypass_shell" => true];
$process = proc_open("php -v", $descriptorspec, $pipes, null, null, $options);
if (is_resource($process)) {
    fclose($pipes[1]);
    proc_close($process);
    echo "BYPASS_SHELL_OK";
} else {
    echo "BYPASS_SHELL_OK";
}
