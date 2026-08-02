<?php
// vybe-test: php/oop_patterns/magic_tostring_representation
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Vector2D {
    public function __construct(private float $x, private float $y) {}
    public function __toString(): string {
        return "({$this->x}, {$this->y})";
    }
    public function add(Vector2D $other): self {
        return new self($this->x + $other->x, $this->y + $other->y);
    }
}
$v1 = new Vector2D(1.0, 2.0);
$v2 = new Vector2D(3.0, 4.0);
echo $v1;
echo $v1->add($v2);
