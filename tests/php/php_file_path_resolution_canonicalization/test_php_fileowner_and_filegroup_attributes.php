<?php
// vybe-test: php/php_file_path_resolution_canonicalization/test_php_fileowner_and_filegroup_attributes
// origin: languages/php/tests/php/test_php_file_path_resolution_canonicalization.rs
// vybe-test-mode: compile

$file = tempnam(sys_get_temp_dir(), "vybe_owner_");
$owner = fileowner($file);
$group = filegroup($file);
echo is_numeric($owner) && is_numeric($group) ? "OWNER_OK" : "FAIL";
unlink($file);
