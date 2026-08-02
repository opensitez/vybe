<?php
// vybe-test: php/oop_patterns/object_comparison_loose_vs_strict
// origin: languages/php/tests/php/test_oop_patterns.rs
// vybe-test-mode: compile

class Point {
    public function __construct(public int $x, public int $y) {}
}
$a = new Point(1, 2);
$b = new Point(1, 2);
$c = $a;
echo ($a == $b)  ? 'loose-eq'  : 'loose-neq';
echo ($a === $b) ? 'strict-eq' : 'strict-neq';
echo ($a === $c) ? 'strict-eq' : 'strict-neq';
