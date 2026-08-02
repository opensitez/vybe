<?php
// vybe-test: php/classes/new_with_args
// origin: languages/php/tests/php/test_classes.rs
// vybe-test-mode: compile

class Point { public $x; public $y; public function __construct($x, $y) { $this->x = $x; $this->y = $y; } } $p = new Point(1, 2);
