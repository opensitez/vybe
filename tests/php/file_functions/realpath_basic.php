<?php
// vybe-test: php/file_functions/realpath_basic
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$real = realpath('/etc/../etc/hostname');
echo is_string($real) || $real === false ? 'ok' : 'fail';
