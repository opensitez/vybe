<?php
// vybe-test: php/php_math_random_engines_php82/test_php82_randomizer_shuffle_bytes_string
// origin: languages/php/tests/php/test_php_math_random_engines_php82.rs
// vybe-test-mode: compile

if (class_exists('Random\Randomizer')) {
    $r = new Random\Randomizer();
    $shuffled = $r->shuffleBytes("abcdef");
    echo strlen($shuffled) === 6 ? "SHUFFLE_BYTES_OK" : "FAIL";
}
