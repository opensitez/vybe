<?php
// vybe-test: php/declare/strict_types_with_interface
// origin: languages/php/tests/php/test_declare.rs
// vybe-test-mode: compile

declare(strict_types=1);
interface Measurable {
    public function measure(): float;
}
class Circle implements Measurable {
    public function __construct(private float $r) {}
    public function measure(): float { return M_PI * $this->r ** 2; }
}
$c = new Circle(3.0);
echo round($c->measure(), 2);
