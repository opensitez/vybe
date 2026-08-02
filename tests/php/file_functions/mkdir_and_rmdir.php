<?php
// vybe-test: php/file_functions/mkdir_and_rmdir
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$path = '/tmp/vybe_test_dir_' . getmypid();
if (!is_dir($path)) {
    $ok = mkdir($path, 0755);
    echo $ok ? 'created' : 'failed';
    $ok2 = rmdir($path);
    echo $ok2 ? ':removed' : ':rmdir failed';
}
