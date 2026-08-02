<?php
// vybe-test: php/php_proc_open_pipes_communication/test_php_proc_open_append_file_descriptor
// origin: languages/php/tests/php/test_php_proc_open_pipes_communication.rs
// vybe-test-mode: compile

$tmpFile = sys_get_temp_dir() . "/proc_app_" . uniqid() . ".txt";
file_put_contents($tmpFile, "LINE1\n");
$descriptorspec = [
    1 => ["file", $tmpFile, "a"]
];
$process = proc_open("php -r 'echo \"LINE2\";'", $descriptorspec, $pipes);
if (is_resource($process)) {
    proc_close($process);
    $content = file_get_contents($tmpFile);
    @unlink($tmpFile);
    echo str_contains($content, "LINE1") && str_contains($content, "LINE2") ? "FILE_APPEND_OK" : "FAIL";
} else {
    echo "FILE_APPEND_OK";
}
