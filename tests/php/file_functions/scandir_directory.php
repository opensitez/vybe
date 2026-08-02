<?php
// vybe-test: php/file_functions/scandir_directory
// origin: languages/php/tests/php/test_file_functions.rs
// vybe-test-mode: compile

$entries = scandir('/tmp');
echo is_array($entries) ? 'array' : 'not array';
echo in_array('.', $entries) ? 'has dot' : 'no dot';
