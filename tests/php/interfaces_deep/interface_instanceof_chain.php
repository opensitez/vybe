<?php
// vybe-test: php/interfaces_deep/interface_instanceof_chain
// origin: languages/php/tests/php/test_interfaces_deep.rs
// vybe-test-mode: compile

interface Shape { public function area(): float; }
interface ColoredShape extends Shape { public function color(): string; }
class RedCircle implements ColoredShape {
    public function __construct(private float $r) {}
    public function area(): float { return M_PI * $this->r ** 2; }
    public function color(): string { return 'red'; }
}
$c = new RedCircle(2.0);
echo ($c instanceof Shape) ? 'is Shape' : 'not Shape';
echo ($c instanceof ColoredShape) ? ':is ColoredShape' : ':not ColoredShape';
