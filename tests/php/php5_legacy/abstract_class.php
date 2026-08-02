<?php
// vybe-test: php/php5_legacy/abstract_class
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

abstract class Shape { abstract public function area(); } class Circle extends Shape { public function area() { return 3.14; } }
