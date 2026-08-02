<?php
// vybe-test: php/file_functions/is_writable_check
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$w = is_writable('/tmp');
echo is_bool($w) ? 'bool' : 'not bool';
