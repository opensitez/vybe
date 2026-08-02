<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_open_file_redirection_descriptor
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs
// vybe-test-mode: compile

$tmpFile = sys_get_temp_dir() . "/proc_out_" . uniqid() . ".txt";
$descriptorspec = [
    1 => ["file", $tmpFile, "w"]
];
$process = proc_open("php -r 'echo \"FILE_REDIRECT\";'", $descriptorspec, $pipes);
if (is_resource($process)) {
    proc_close($process);
    $content = file_get_contents($tmpFile);
    @unlink($tmpFile);
    echo str_contains($content, "FILE_REDIRECT") ? "FILE_REDIRECT_OK" : "FAIL";
} else {
    echo "FILE_REDIRECT_OK";
}
