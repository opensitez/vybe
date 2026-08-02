<?php
// vybe-test: php/php_streams_file_system_wrapper_ops/test_php_mkdir_rmdir_recursive_directory
// origin: languages/php/tests/php/test_php_streams_file_system_wrapper_ops.rs
// vybe-test-mode: compile

$dir = sys_get_temp_dir() . "/nested/dir/test";
if (!is_dir($dir)) {
    mkdir($dir, 0755, recursive: true);
}
echo is_dir($dir) ? "DIR_CREATED" : "DIR_FAILED";
rmdir($dir);
