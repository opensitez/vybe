<?php
// vybe-test: php/traits_deep/trait_late_binding_in_static
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Registry {
    private static array $items = [];
    public static function register(string $key, mixed $val): void { static::$items[$key] = $val; }
    public static function get(string $key): mixed { return static::$items[$key] ?? null; }
    public static function all(): array { return static::$items; }
}
class ServiceContainer { use Registry; }
ServiceContainer::register('db', 'sqlite');
ServiceContainer::register('cache', 'redis');
echo count(ServiceContainer::all());
