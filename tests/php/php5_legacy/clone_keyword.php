<?php
// vybe-test: php/php5_legacy/clone_keyword
// origin: languages/php/tests/php/test_php5_legacy.rs
// vybe-test-mode: compile

class A { public $x = 1; } $a = new A(); $b = clone $a;
