<?php
// vybe-test: php/namespaces/namespace_abstract_class
// origin: languages/php/tests/php/test_namespaces.rs
// vybe-test-mode: compile

namespace Base;
abstract class Shape {
    abstract public function area(): float;
    public function describe(): string { return "area=" . $this->area(); }
}

namespace Shapes;
use Base\Shape;
class Circle extends Shape {
    public function __construct(private float $r) {}
    public function area(): float { return M_PI * $this->r ** 2; }
}
$c = new Circle(2.0);
echo round($c->area(), 4);
