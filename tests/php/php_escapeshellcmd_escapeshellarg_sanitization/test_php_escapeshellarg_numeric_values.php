<?php
// vybe-test: php/php_escapeshellcmd_escapeshellarg_sanitization/test_php_escapeshellarg_numeric_values
// origin: languages/php/tests/php/test_php_escapeshellcmd_escapeshellarg_sanitization.rs
// vybe-test-mode: compile

$arg = 12345;
$escaped = escapeshellarg($arg);
echo $escaped === "'12345'" || $escaped === '"12345"' ? "NUMERIC_ARG_ESCAPED_OK" : "FAIL";
