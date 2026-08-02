<?php
// vybe-test: php/variable_variables/call_user_func_method
// origin: languages/php/tests/php/test_variable_variables.rs
// vybe-test-mode: compile

class Calc {
    public function add(int $a, int $b): int { return $a + $b; }
}
$c = new Calc();
echo call_user_func([$c, 'add'], 10, 32);
