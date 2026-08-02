<?php
// vybe-test: php/oop/abstract_class
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

abstract class Shape {
    abstract public function area(): float;
    public function describe() { return 'Shape with area ' . $this->area(); }
}
class Circle extends Shape {
    public $radius;
    public function __construct($r) { $this->radius = $r; }
    public function area(): float { return 3.14159 * $this->radius * $this->radius; }
}
$c = new Circle(5);
echo $c->describe();
