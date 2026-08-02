<?php
// vybe-test: php/oop/class_methods
// origin: languages/php/tests/php/test_oop.rs
// vybe-test-mode: compile

class Calc { public function add($a, $b) { return $a + $b; } public function sub($a, $b) { return $a - $b; } } $c = new Calc(); echo $c->add(3, 2);
