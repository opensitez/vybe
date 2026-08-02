<?php
// vybe-test: php/file_functions/is_dir_check
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

echo is_dir('/etc') ? 'dir' : 'not dir';
echo is_dir('/etc/hostname') ? 'dir' : 'not dir';
