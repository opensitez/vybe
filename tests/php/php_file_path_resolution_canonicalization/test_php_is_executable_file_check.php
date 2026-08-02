<?php
// vybe-test: php/php_file_path_resolution_canonicalization/test_php_is_executable_file_check
// origin: languages/php/tests/php/test_php_file_path_resolution_canonicalization.rs
// vybe-test-mode: compile

$file = tempnam(sys_get_temp_dir(), "vybe_exec_");
echo is_executable($file) ? "EXEC" : "NON_EXEC";
unlink($file);
