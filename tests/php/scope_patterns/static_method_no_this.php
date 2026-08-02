<?php
// vybe-test: php/scope_patterns/static_method_no_this
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

class MathHelper {
    public static function square(int $n): int { return $n * $n; }
}
echo MathHelper::square(9);
