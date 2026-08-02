<?php
// vybe-test: php/php_array_chunk_combine_count_values/test_php_array_reduce_string_concatenation
// origin: languages/php/tests/php/test_php_array_chunk_combine_count_values.rs
// vybe-test-mode: compile

$words = ["PHP", "Is", "Great"];
$sentence = array_reduce($words, fn($carry, $w) => $carry === "" ? $w : "$carry $w", "");
echo $sentence;
