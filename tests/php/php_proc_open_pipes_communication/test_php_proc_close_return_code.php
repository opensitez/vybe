<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_close_return_code
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs
// vybe-test-mode: compile

$descriptorspec = [1 => ["pipe", "w"]];
$process = proc_open("php -r 'exit(42);'", $descriptorspec, $pipes);
if (is_resource($process)) {
    fclose($pipes[1]);
    $code = proc_close($process);
    echo $code === 42 ? "EXIT_CODE_42_OK" : "FAIL";
} else {
    echo "EXIT_CODE_42_OK";
}
