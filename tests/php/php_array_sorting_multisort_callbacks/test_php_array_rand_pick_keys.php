<?php
// vybe-test: php/php_array_sorting_multisort_callbacks/test_php_array_rand_pick_keys
// origin: languages/php/tests/php/test_php_array_sorting_multisort_callbacks.rs
// vybe-test-mode: compile

$input = ["Neo", "Morpheus", "Trinity", "Cypher", "Tank"];
$randKey = array_rand($input);
echo $input[$randKey];
