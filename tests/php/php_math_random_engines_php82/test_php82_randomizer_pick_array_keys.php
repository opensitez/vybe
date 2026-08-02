<?php
// vybe-test: php/php_math_random_engines_php82/test_php82_randomizer_pick_array_keys
// origin: languages/php/tests/php/test_php_math_random_engines_php82.rs
// vybe-test-mode: compile

if (class_exists('Random\Randomizer')) {
    $r = new Random\Randomizer();
    $input = ["a" => 1, "b" => 2, "c" => 3, "d" => 4];
    $keys = $r->pickArrayKeys($input, 2);
    echo count($keys) === 2 ? "PICK_KEYS_OK" : "FAIL";
}
