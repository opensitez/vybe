<?php
// vybe-test: php/php5_legacy/constructor
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

class A { public $x; public function __construct($x) { $this->x = $x; } } $a = new A(42); echo $a->x;
