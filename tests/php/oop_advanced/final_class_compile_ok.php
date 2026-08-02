<?php
// vybe-test: php/oop_advanced/final_class_compile_ok
// origin: languages/php/tests/php/test_oop_advanced.rs
// vybe-test-mode: compile

final class Singleton {
    private static ?self $instance = null;
    private function __construct(public readonly string $id) {}
    public static function getInstance(): self {
        if (self::$instance === null) {
            self::$instance = new self("main");
        }
        return self::$instance;
    }
}
$s = Singleton::getInstance();
echo $s->id, "\n";
