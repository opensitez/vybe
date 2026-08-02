<?php
// vybe-test: php/advanced_closures/static_closure_no_this_binding
// origin: languages/php/tests/php/test_advanced_closures.rs
// vybe-test-mode: compile

class Widget {
    public static function getTransformer(): Closure {
        return static function(int $n): int { return $n ** 2; };
    }
}
$fn = Widget::getTransformer();
echo $fn(6);
