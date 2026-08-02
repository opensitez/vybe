<?php
// vybe-test: php/closures_advanced/closure_bind_static_context
// origin: languages/php/tests/php/test_closures_advanced.rs
// vybe-test-mode: compile

class Registry { private static array $items = []; }
$add = Closure::bind(
    static function(string $k, mixed $v) { static::$items[$k] = $v; },
    null,
    Registry::class
);
$add('key', 'value');
$get = Closure::bind(
    static function(string $k) { return static::$items[$k] ?? null; },
    null,
    Registry::class
);
echo $get('key');
