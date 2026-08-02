<?php
// vybe-test: php/functional_style/array_map_returning_objects
// origin: languages/php/tests/php/test_functional_style.rs
// vybe-test-mode: compile

class Point {
    public function __construct(public int $x, public int $y) {}
}
$coords = [[1, 2], [3, 4], [5, 6]];
$points = array_map(fn($c) => new Point($c[0], $c[1]), $coords);
echo $points[1]->x . ',' . $points[1]->y;
echo count($points);
