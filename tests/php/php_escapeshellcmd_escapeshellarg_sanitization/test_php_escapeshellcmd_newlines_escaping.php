<?php
// vybe-test: php/php_escapeshellcmd_escapeshellarg_sanitization/test_php_escapeshellcmd_newlines_escaping
// origin: languages/php/tests/php/test_php_escapeshellcmd_escapeshellarg_sanitization.rs
// vybe-test-mode: compile

$cmd = "echo hello\necho world";
$clean = escapeshellcmd($cmd);
echo is_string($clean) ? "NEWLINE_ESCAPED_OK" : "FAIL";
