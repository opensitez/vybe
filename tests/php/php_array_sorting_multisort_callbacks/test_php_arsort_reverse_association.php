<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_arsort_reverse_association
// origin: languages/php/tests/php/test_php_array_sorting_multisort_callbacks.rs
// vybe-test-mode: compile

$scores = ["Alice" => 90, "Bob" => 95, "Charlie" => 85];
arsort($scores);
$top = array_key_first($scores);
echo "$top=" . $scores[$top];
