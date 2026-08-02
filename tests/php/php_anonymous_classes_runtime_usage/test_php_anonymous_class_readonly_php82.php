<?php
// vybe-test: php/php_anonymous_classes_runtime_usage/test_php_anonymous_class_readonly_php82
// origin: languages/php/tests/php/test_php_anonymous_classes_runtime_usage.rs
// vybe-test-mode: compile

$immutable = new readonly class(10, 20) {
    public function __construct(public int $x, public int $y) {}
};

echo "{$immutable->x}, {$immutable->y}";
