<?php
// vybe-test: php/php_traits_composition_conflict_resolution/test_php_trait_static_methods_and_properties
// origin: languages/php/tests/php/test_php_traits_composition_conflict_resolution.rs
// vybe-test-mode: compile

trait SingletonTrait {
    private static ?self $instance = null;
    public static function getInstance(): self {
        return self::$instance ??= new self();
    }
}

class AppConfig {
    use SingletonTrait;
}

$app = AppConfig::getInstance();
echo get_class($app);
