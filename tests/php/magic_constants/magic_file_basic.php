<?php
// vybe-test: php/magic_constants/magic_file_basic
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

$file = __FILE__;
echo is_string($file) ? 'is string' : 'not string';
