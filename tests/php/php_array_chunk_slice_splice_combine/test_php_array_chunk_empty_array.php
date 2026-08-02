<?php
// vybe-test: php/php_array_chunk_slice_splice_combine/test_php_array_chunk_empty_array
// origin: languages/php/tests/php/test_php_array_chunk_slice_splice_combine.rs
// vybe-test-mode: compile

$c = array_chunk([], 3);
echo count($c) === 0 ? "EMPTY_CHUNK_OK" : "FAIL";
