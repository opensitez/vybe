<?php
// vybe-test: php/oop_advanced/final_method_in_hierarchy
// origin: languages/php/tests/php/test_oop_advanced.rs
// vybe-test-mode: compile

class Base {
    final public function identity(): string {
        return static::class;
    }
    public function greeting(): string {
        return "Hello from " . $this->identity();
    }
}
class Child extends Base {
    public function greeting(): string {
        return parent::greeting() . " (child)";
    }
}
$c = new Child();
echo $c->greeting(), "\n";
