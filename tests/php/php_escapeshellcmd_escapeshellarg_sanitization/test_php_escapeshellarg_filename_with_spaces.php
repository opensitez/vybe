<?php
// vybe-test: php/php_escapeshellcmd_escapeshellarg_sanitization/test_php_escapeshellarg_filename_with_spaces
// origin: languages/php/tests/php/test_php_escapeshellcmd_escapeshellarg_sanitization.rs
// vybe-test-mode: compile

$filename = "my file name.pdf";
$escaped = escapeshellarg($filename);
echo $escaped === "'my file name.pdf'" || $escaped === '"my file name.pdf"' ? "SPACES_FILENAME_OK" : "FAIL";
