<?php
// vybe-test: php/file_functions/is_readable_check
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$r = is_readable('/etc/hostname');
echo is_bool($r) ? 'bool' : 'not bool';
