<?php
// vybe-test: php/phase2/ctor_promotion_full
// origin: languages/php/tests/php/test_phase2.rs
// vybe-test-mode: compile

class Point {
    public function __construct(
        public float $x,
        public float $y,
        public float $z = 0.0
    ) {}
}
$p = new Point(1.0, 2.0);
