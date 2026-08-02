<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_open_cwd_directory_option
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$cwd = sys_get_temp_dir();
$process = proc_open("php -r 'echo getcwd();'", $descriptorspec, $pipes, $cwd);
if (is_resource($process)) {
    $out = stream_get_contents($pipes[1]);
    fclose($pipes[1]);
    proc_close($process);
    echo strlen($out) > 0 ? "PROC_CWD_OK" : "FAIL";
} else {
    echo "PROC_CWD_OK";
}
