<?php
// vybe-test: php/oop/ctor_promotion
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Point { public function __construct(public float $x, public float $y) {} } $p = new Point(1.0, 2.0);
