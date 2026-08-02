<?php
// vybe-test: php/magic_constants/magic_dir_vs_file
// origin: languages/php/tests/php/test_magic_constants.rs
// vybe-test-mode: compile

$dir  = __DIR__;
$file = __FILE__;
echo strlen($dir) <= strlen($file) ? 'dir shorter or equal' : 'dir longer';
