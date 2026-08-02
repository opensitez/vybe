<?php
// vybe-test: php/scope_patterns/class_method_uses_this
// origin: languages/php/tests/php/test_scope_patterns.rs
// vybe-test-mode: compile

class Box {
    private int $value;
    public function __construct(int $v) { $this->value = $v; }
    public function get(): int { return $this->value; }
}
echo (new Box(77))->get();
