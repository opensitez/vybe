<?php
// vybe-test: php/php_escapeshellcmd_escapeshellarg_sanitization/test_php_escapeshellarg_empty_string
// origin: languages/php/tests/php/test_php_escapeshellcmd_escapeshellarg_sanitization.rs
// vybe-test-mode: compile

$escaped = escapeshellarg("");
echo $escaped === "''" || $escaped === '""' ? "EMPTY_SHELL_ARG_OK" : "FAIL";
