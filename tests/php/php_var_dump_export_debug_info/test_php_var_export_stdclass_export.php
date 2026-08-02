<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_var_export_stdclass_export
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs
// vybe-test-mode: compile

$obj = new stdClass();
$obj->title = "Test";
$exported = var_export($obj, return: true);
echo str_contains($exported, "stdClass::__set_state") ? "SET_STATE_EXPORT" : "ANON_EXPORT";
