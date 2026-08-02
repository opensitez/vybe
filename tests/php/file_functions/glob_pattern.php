<?php
// vybe-test: php/file_functions/glob_pattern
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$files = glob('/tmp/*.txt');
echo is_array($files) ? 'array' : 'not array';
