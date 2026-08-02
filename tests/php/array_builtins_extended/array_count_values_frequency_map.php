<?php
// vybe-test: php/array_builtins_extended/array_count_values_frequency_map
// origin: languages/php/tests/php/test_array_builtins_extended.rs
// vybe-test-mode: compile

$a = ["red", "blue", "red", "green", "blue", "red"];
$freq = array_count_values($a);
echo $freq["red"];
echo $freq["blue"];
echo $freq["green"];
