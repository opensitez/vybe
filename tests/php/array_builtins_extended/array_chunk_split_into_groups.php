<?php
// vybe-test: php/array_builtins_extended/array_chunk_split_into_groups
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = range(1, 7);
$chunks = array_chunk($a, 3);
echo count($chunks);
echo count($chunks[0]);
echo count($chunks[2]);
