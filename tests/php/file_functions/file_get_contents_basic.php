<?php
// vybe-test: php/file_functions/file_get_contents_basic
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$contents = file_get_contents('/etc/hostname');
echo is_string($contents) ? 'string' : 'not string';
