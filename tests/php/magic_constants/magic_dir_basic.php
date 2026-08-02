<?php
// vybe-test: php/magic_constants/magic_dir_basic
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

$dir = __DIR__;
echo is_string($dir) ? 'is string' : 'not string';
