<?php
// vybe-test: php/php_array_chunk_slice_splice_combine/test_php_array_splice_insert_without_deleting
// origin: languages/php/tests/php/test_php_array_chunk_slice_splice_combine.rs
// vybe-test-mode: compile

$a = ["first", "last"];
array_splice($a, 1, 0, ["middle"]);
echo implode(",", $a);
