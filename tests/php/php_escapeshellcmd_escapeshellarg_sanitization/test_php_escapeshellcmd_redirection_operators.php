<?php
// vybe-test: php/php_escapeshellcmd_escapeshellarg_sanitization/test_php_escapeshellcmd_redirection_operators
// origin: languages/php/tests/php/test_php_escapeshellcmd_escapeshellarg_sanitization.rs
// vybe-test-mode: compile

$cmd = "cat < input.txt > output.txt 2>&1";
$clean = escapeshellcmd($cmd);
echo str_contains($clean, "\\<") && str_contains($clean, "\\>") ? "REDIRECTION_ESCAPED_OK" : "FAIL";
