<?php
// vybe-test: php/php_escapeshellcmd_escapeshellarg_sanitization/test_php_escapeshellcmd_escapes_dollar_and_backticks
// origin: languages/php/tests/php/test_php_escapeshellcmd_escapeshellarg_sanitization.rs
// vybe-test-mode: compile

$input = "echo \$VAR `whoami` (sub)";
$clean = escapeshellcmd($input);
echo str_contains($clean, "\\$") && str_contains($clean, "\\`") ? "METACHARS_ESCAPED_OK" : "FAIL";
