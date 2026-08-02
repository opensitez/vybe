<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_uasort_custom_association_sort
// origin: languages/php/tests/php/test_php_array_sorting_multisort_callbacks.rs
// vybe-test-mode: compile

$data = ["a" => 4, "b" => 2, "c" => 8];
uasort($data, fn($v1, $v2) => $v1 <=> $v2);
echo array_key_first($data);
