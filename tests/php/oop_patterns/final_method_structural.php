<?php
// vybe-test: php/oop_patterns/final_method_structural
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Base {
    final public function identity(): string { return static::class; }
    public function greet(): string { return 'hello from ' . $this->identity(); }
}
class Child extends Base {
    public function greet(): string { return 'child greet: ' . $this->identity(); }
}
$b = new Base();
$c = new Child();
echo $b->greet();
echo $c->greet();
