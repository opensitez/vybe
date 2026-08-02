<?php
// vybe-test: php/array_builtins_extended/natsort_natural_string_order
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$files = ["file10.txt", "file2.txt", "file1.txt", "file20.txt"];
natsort($files);
echo implode(",", $files);
