<?php
// vybe-test: php/class_inspection/get_object_vars_public_properties
// origin: languages/php/tests/php/test_class_inspection.rs
// vybe-test-mode: compile

class Point {
    public int $x;
    public int $y;
    public function __construct(int $x, int $y) {
        $this->x = $x;
        $this->y = $y;
    }
}
$p = new Point(3, 7);
$vars = get_object_vars($p);
echo $vars['x'];
echo $vars['y'];
