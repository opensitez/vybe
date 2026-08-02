<?php
// vybe-test: php/php_streams_file_system_wrapper_ops/test_php_glob_file_search_pattern
// origin: languages/php/tests/php/test_php_streams_file_system_wrapper_ops.rs
// vybe-test-mode: compile

$files = glob(sys_get_temp_dir() . "/*");
echo is_array($files) ? "ARRAY" : "FALSE";
