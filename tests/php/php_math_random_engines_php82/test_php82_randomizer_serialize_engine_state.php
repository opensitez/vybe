<?php
// vybe-test: php/php_math_random_engines_php82/test_php82_randomizer_serialize_engine_state
// origin: languages/php/tests/php/test_php_math_random_engines_php82.rs
// vybe-test-mode: compile

if (class_exists('Random\Engine\Xoshiro256StarStar')) {
    $e = new Random\Engine\Xoshiro256StarStar(99);
    $serialized = serialize($e);
    $restored = unserialize($serialized);
    echo get_class($restored);
}
