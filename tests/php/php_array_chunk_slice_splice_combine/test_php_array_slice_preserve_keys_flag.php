<?php
// vybe-test: php/php_array_chunk_slice_splice_combine/test_php_array_slice_preserve_keys_flag
// origin: languages/php/tests/php/test_php_array_chunk_slice_splice_combine.rs
// vybe-test-mode: compile

$arr = [10 => "ten", 20 => "twenty", 30 => "thirty"];
$sliced = array_slice($arr, 1, 2, true);
echo isset($sliced[20]) ? "KEY_PRESERVED" : "FAIL";
