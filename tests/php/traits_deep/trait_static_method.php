<?php
// vybe-test: php/traits_deep/trait_static_method
// origin: languages/php/tests/php/test_traits_deep.rs
// vybe-test-mode: compile

trait Singleton {
    private static ?self $instance = null;
    public static function getInstance(): static {
        if (static::$instance === null) { static::$instance = new static(); }
        return static::$instance;
    }
}
class Config { use Singleton; public string $env = 'production'; }
$c1 = Config::getInstance();
$c2 = Config::getInstance();
$c1->env = 'staging';
echo $c2->env;
