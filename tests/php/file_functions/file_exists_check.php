<?php
// vybe-test: php/file_functions/file_exists_check
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$exists = file_exists('/etc/hostname');
echo is_bool($exists) ? 'bool' : 'not bool';
