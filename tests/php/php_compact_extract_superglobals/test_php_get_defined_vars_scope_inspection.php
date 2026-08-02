<?php
// vybe-test: php/php_compact_extract_superglobals/test_php_get_defined_vars_scope_inspection
// origin: languages/php/tests/php/test_php_compact_extract_superglobals.rs
// vybe-test-mode: compile

$x = 10;
$y = "hello";
$vars = get_defined_vars();
echo isset($vars["x"]) && isset($vars["y"]) ? "VARS_FOUND" : "FAIL";
