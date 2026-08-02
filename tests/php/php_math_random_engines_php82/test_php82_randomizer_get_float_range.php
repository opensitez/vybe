<?php
// vybe-test: php/php_math_random_engines_php82/test_php82_randomizer_get_float_range
// origin: languages/php/tests/php/test_php_math_random_engines_php82.rs
// vybe-test-mode: compile

if (method_exists('Random\Randomizer', 'getFloat')) {
    $r = new Random\Randomizer();
    $f = $r->getFloat(0.0, 1.0);
    echo ($f >= 0.0 && $f <= 1.0) ? "FLOAT_RANGE_OK" : "FAIL";
}
