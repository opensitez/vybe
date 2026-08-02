<?php
// vybe-test: php/named_arguments/named_args_with_abstract_method
// origin: languages/php/tests/php/test_named_arguments.rs
// vybe-test-mode: compile

abstract class Shape {
    abstract public function area(float $scale = 1.0): float;
}
class Circle extends Shape {
    public function __construct(private float $radius) {}
    public function area(float $scale = 1.0): float {
        return M_PI * $this->radius ** 2 * $scale;
    }
}
$c = new Circle(radius: 5.0);
echo round($c->area(scale: 2.0), 2);
