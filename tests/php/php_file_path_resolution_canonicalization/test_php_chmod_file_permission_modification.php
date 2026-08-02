<?php
// vybe-test: php/php_file_path_resolution_canonicalization/test_php_chmod_file_permission_modification
// origin: languages/php/tests/php/test_php_file_path_resolution_canonicalization.rs
// vybe-test-mode: compile

$file = tempnam(sys_get_temp_dir(), "vybe_chmod_");
chmod($file, 0644);
$perms = fileperms($file);
echo is_numeric($perms) ? "PERMS_OK" : "FAIL";
unlink($file);
