<?php
// vybe-test: php/variable_variables/call_user_func_static_method
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class MathUtil {
    public static function square(int $n): int { return $n * $n; }
}
echo call_user_func(['MathUtil', 'square'], 9);
