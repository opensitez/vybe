<?php
// vybe-test: php/php84/readonly_class_deep
// origin: languages/php/tests/php/test_php84.rs
// vybe-test-mode: compile

readonly class Coordinate {
    public function __construct(
        public float $lat,
        public float $lon
    ) {}
    public function distanceTo(Coordinate $other): float {
        return sqrt(($this->lat - $other->lat)**2 + ($this->lon - $other->lon)**2);
    }
}
$a = new Coordinate(0.0, 0.0);
$b = new Coordinate(3.0, 4.0);
echo $b->distanceTo($a);
