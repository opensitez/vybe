<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_var_export_resource_behavior
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs
// vybe-test-mode: compile

$fp = fopen("php://memory", "r");
$exp = var_export($fp, return: true);
fclose($fp);
echo str_contains($exp, "NULL") ? "RESOURCE_EXPORT_NULL" : "EXPORT_OK";
