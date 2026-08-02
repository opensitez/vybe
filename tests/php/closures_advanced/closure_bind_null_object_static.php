<?php
// vybe-test: php/closures_advanced/closure_bind_null_object_static
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Config {
    private static string $env = 'production';
    public static function getEnv(): string { return static::$env; }
}
$reader = Closure::bind(
    static function() { return static::$env; },
    null,
    Config::class
);
echo $reader();
