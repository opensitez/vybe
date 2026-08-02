<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_krsort_key_reverse_order
// origin: languages/php/tests/php/test_php_array_sorting_multisort_callbacks.rs
// vybe-test-mode: compile

$arr = [1 => "a", 3 => "c", 2 => "b"];
krsort($arr);
echo implode(",", array_keys($arr));
