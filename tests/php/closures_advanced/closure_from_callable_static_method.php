<?php
// vybe-test: php/closures_advanced/closure_from_callable_static_method
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Util {
    public static function triple(int $n): int { return $n * 3; }
}
$fn = Closure::fromCallable(['Util', 'triple']);
echo $fn(7);
