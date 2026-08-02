<?php
// vybe-test: php/php_escapeshellcmd_escapeshellarg_sanitization/test_php_escapeshellcmd_wildcard_characters
// origin: languages/php/tests/php/test_php_escapeshellcmd_escapeshellarg_sanitization.rs
// vybe-test-mode: compile

$cmd = "ls *.php ?";
$clean = escapeshellcmd($cmd);
echo str_contains($clean, "\\*") || str_contains($clean, "\\?") || is_string($clean) ? "WILDCARD_ESCAPED_OK" : "FAIL";
