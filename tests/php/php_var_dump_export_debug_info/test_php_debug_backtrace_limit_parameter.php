<?php
// vybe-test: php/php_var_dump_export_debug_info/test_php_debug_backtrace_limit_parameter
// origin: languages/php/tests/php/test_php_var_dump_export_debug_info.rs
// vybe-test-mode: compile

function level3() { return debug_backtrace(limit: 2); }
function level2() { return level3(); }
function level1() { return level2(); }

$frames = level1();
echo count($frames) <= 2 ? "LIMIT_2_OK" : "LIMIT_EXCEEDED";
