<?php
// vybe-test: php/php_math_random_engines_php82/test_php82_random_engine_pcg63810xy256_seeding
// origin: languages/php/tests/php/test_php_math_random_engines_php82.rs
// vybe-test-mode: compile

if (class_exists('Random\Engine\Pcg63810XY256')) {
    $engine = new Random\Engine\Pcg63810XY256(123);
    $v1 = $engine->generate();
    echo is_string($v1) && strlen($v1) > 0 ? "PCG_GENERATE_OK" : "FAIL";
}
