<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_shuffle_randomize_array
// origin: languages/php/tests/php/test_php_array_sorting_multisort_callbacks.rs
// vybe-test-mode: compile

$numbers = range(1, 10);
shuffle($numbers);
echo count($numbers);
