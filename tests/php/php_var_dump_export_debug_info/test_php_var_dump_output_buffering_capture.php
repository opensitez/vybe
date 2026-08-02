<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_var_dump_output_buffering_capture
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs
// vybe-test-mode: compile

ob_start();
$val = ["hello", 123, true, null];
var_dump($val);
$dump = ob_get_clean();

echo str_contains($dump, "string(5)") && str_contains($dump, "int(123)") ? "VAR_DUMP_OK" : "FAIL";
