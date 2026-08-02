<?php
// vybe-test: php/scope_patterns/self_vs_static_in_static_context
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

class Base {
    protected static string $label = 'Base';
    public static function selfLabel(): string  { return self::$label; }
    public static function lateLabel(): string  { return static::$label; }
}
class Child extends Base {
    protected static string $label = 'Child';
}
echo Base::selfLabel();
echo Child::lateLabel();
