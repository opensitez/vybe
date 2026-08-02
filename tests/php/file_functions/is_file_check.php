<?php
// vybe-test: php/file_functions/is_file_check
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

echo is_file('/etc/hostname') ? 'file' : 'not file';
echo is_file('/etc') ? 'file' : 'not file';
