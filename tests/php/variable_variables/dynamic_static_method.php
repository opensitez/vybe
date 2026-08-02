<?php
// vybe-test: php/variable_variables/dynamic_static_method
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Factory {
    public static function makeString(): string { return 'str'; }
    public static function makeInt(): int { return 42; }
}
$method = 'makeString';
echo Factory::$method();
