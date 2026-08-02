<?php
// vybe-test: php/php_oop_late_static_binding_self_static/test_php_late_static_binding_in_singleton
// origin: languages/php/tests/php/test_php_oop_late_static_binding_self_static.rs
// vybe-test-mode: compile

abstract class Singleton {
    private static array $instances = [];
    public static function getInstance(): static {
        $cls = static::class;
        return self::$instances[$cls] ??= new static();
    }
}

class AppRegistry extends Singleton {}
$reg = AppRegistry::getInstance();
echo get_class($reg);
