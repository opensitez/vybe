<?php
// vybe-test: php/php5_legacy/deep_property
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

class A { public $b; } class B { public $c = 42; } $a = new A(); $a->b = new B(); echo $a->b->c;
