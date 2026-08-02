<?php
// vybe-test: php/file_functions/file_read_lines
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$lines = file('/etc/hostname', FILE_IGNORE_NEW_LINES);
echo is_array($lines) ? 'array' : 'not array';
