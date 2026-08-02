<?php
// vybe-test: php/php_math_random_engines_php82/test_php82_random_engine_user_custom_implementation
// origin: languages/php/tests/php/test_php_math_random_engines_php82.rs
// vybe-test-mode: compile

if (interface_exists('Random\Engine')) {
    class ConstantEngine implements Random\Engine {
        public function generate(): string {
            return "\x00\x00\x00\x00\x00\x00\x00\x00";
        }
    }
    $r = new Random\Randomizer(new ConstantEngine());
    echo "Custom engine instantiated";
}
